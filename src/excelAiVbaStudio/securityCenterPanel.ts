import { randomBytes } from 'crypto';
import * as path from 'path';
import * as vscode from 'vscode';
import {
	EnterpriseSecurityReport,
	formatEnterpriseSecurityReport,
	OfficeSecurityService,
	OfficeSecuritySource,
	OfficeSecurityStatus
} from './officeSecurity';

type OpenExcelCallback = (workbookUri: vscode.Uri) => Promise<boolean>;

interface SecurityCenterMessage {
	type?: unknown;
}

const MICROSOFT_CLOUD_POLICY_DOCS_URL = vscode.Uri.parse(
	'https://learn.microsoft.com/en-us/microsoft-365-apps/admin-center/overview-cloud-policy'
);
const MICROSOFT_365_APPS_ADMIN_URL = vscode.Uri.parse('https://config.office.com');

const STATUS_LABELS: Record<OfficeSecurityStatus, string> = {
	protected: 'Protégé',
	blocked: 'Bloqué',
	prompt: 'Confirmation requise',
	allowed: 'Autorisé',
	managed: 'Géré',
	warning: 'Attention',
	unknown: 'À confirmer',
	notApplicable: 'Sans objet'
};

const SOURCE_LABELS: Record<OfficeSecuritySource, string> = {
	machinePolicy: 'Stratégie ordinateur',
	userPolicy: 'Stratégie utilisateur',
	cloudPolicy: 'Cloud Policy Microsoft 365',
	userPreference: 'Préférence utilisateur',
	machinePreference: 'Préférence ordinateur'
};

const LEVEL_LABELS: Record<EnterpriseSecurityReport['level'], string> = {
	restricted: 'Restreint',
	managed: 'Géré par l’organisation',
	standard: 'Standard',
	unknown: 'À confirmer dans Excel'
};

const ZONE_LABELS: Record<number, string> = {
	0: 'Ordinateur local',
	1: 'Intranet local',
	2: 'Sites de confiance',
	3: 'Internet',
	4: 'Sites sensibles'
};

function escapeHtml(value: unknown): string {
	return String(value ?? '').replace(/[&<>"']/g, character => {
		switch (character) {
			case '&':
				return '&amp;';
			case '<':
				return '&lt;';
			case '>':
				return '&gt;';
			case '"':
				return '&quot;';
			default:
				return '&#39;';
		}
	});
}

function securityStatus(status: OfficeSecurityStatus): string {
	return `status-${status}`;
}

function renderStatus(status: OfficeSecurityStatus): string {
	return `<span class="status ${securityStatus(status)}">${escapeHtml(
		STATUS_LABELS[status]
	)}</span>`;
}

function renderSource(source: OfficeSecuritySource): string {
	return SOURCE_LABELS[source] || source;
}

function renderSettingValue(value: string | number | boolean | null): string {
	if (value === null) {
		return 'Non défini';
	}
	if (typeof value === 'boolean') {
		return value ? 'Oui' : 'Non';
	}
	return String(value);
}

function renderReport(report: EnterpriseSecurityReport): string {
	const { probe } = report;
	const zone = probe.workbook.zoneStatus === 'unreadable'
		? 'Origine illisible'
		: probe.workbook.zoneStatus === 'unsupported'
			? 'Origine non prise en charge'
			: probe.workbook.zoneStatus === 'absent'
				? 'Aucune marque d’origine'
				: probe.workbook.zoneId === null
					? 'Flux présent, zone inconnue'
			: `${ZONE_LABELS[probe.workbook.zoneId] || 'Zone inconnue'} (zone ${
					probe.workbook.zoneId
			  })`;
	const managedCount = probe.office.settings.filter(setting => setting.managed)
		.length;
	const labelCount = probe.workbook.sensitivityLabelIds.length;
	const findingRows = report.findings
		.map(
			finding => `
				<article class="rule-card">
					<div class="rule-heading">
						<h3>${escapeHtml(finding.title)}</h3>
						${renderStatus(finding.status)}
					</div>
					<p>${escapeHtml(finding.detail)}</p>
					<p class="impact"><strong>Impact :</strong> ${escapeHtml(
						finding.impact
					)}</p>
					<div class="source-line">
						<span>${escapeHtml(finding.source)}</span>
						${
							finding.managed
								? '<span class="managed-badge">Verrouillé par l’organisation</span>'
								: ''
						}
					</div>
				</article>`
		)
		.join('');
	const capabilityRows = report.capabilities
		.map(
			capability => `
				<tr>
					<td><strong>${escapeHtml(capability.title)}</strong></td>
					<td>${renderStatus(capability.status)}</td>
					<td>${escapeHtml(capability.detail)}</td>
				</tr>`
		)
		.join('');
	const settingRows = probe.office.settings
		.map(
			setting => `
				<tr>
					<td><code>${escapeHtml(setting.id)}</code></td>
					<td>${escapeHtml(renderSettingValue(setting.value))}</td>
					<td>${escapeHtml(renderSource(setting.source))}${
						setting.managed
							? '<br><span class="managed-badge">Géré</span>'
							: ''
					}</td>
					<td><details><summary>Emplacement</summary><code>${escapeHtml(
						setting.registryPath
					)}\\${escapeHtml(setting.name)}</code>${
						setting.registryView
							? `<br><span>Vue ${escapeHtml(setting.registryView)} bits</span>`
							: ''
					}</details></td>
				</tr>`
		)
		.join('');
	const trustedLocationRows = probe.office.trustedLocations
		.map(
			location => `
				<tr>
					<td><code>${escapeHtml(location.path)}</code></td>
					<td>${location.allowSubfolders ? 'Oui' : 'Non'}</td>
					<td>${escapeHtml(renderSource(location.source))}${
						location.managed
							? '<br><span class="managed-badge">Géré</span>'
							: ''
					}</td>
				</tr>`
		)
		.join('');
	const sensitivityLabels =
		labelCount > 0
			? `<ul>${probe.workbook.sensitivityLabelIds
					.map(label => `<li><code>${escapeHtml(label)}</code></li>`)
					.join('')}</ul>`
			: '<p class="empty">Aucune métadonnée d’étiquette Microsoft Purview lisible.</p>';

	return `
		<section class="hero level-${escapeHtml(report.level)}">
			<div>
				<p class="eyebrow">Niveau de sécurité détecté</p>
				<h1>${escapeHtml(LEVEL_LABELS[report.level])}</h1>
				<p class="summary">${escapeHtml(report.summary)}</p>
			</div>
			<div class="level-mark" aria-hidden="true">${
				report.level === 'restricted'
					? '!'
					: report.level === 'managed'
						? 'G'
						: report.level === 'unknown'
							? '?'
							: '✓'
			}</div>
		</section>

		<section class="workbook-card">
			<div>
				<h2>${escapeHtml(probe.workbook.name)}</h2>
				<p class="path">${escapeHtml(probe.workbook.path)}</p>
			</div>
			<div class="facts" role="list">
				<div role="listitem"><span>Origine Windows</span><strong>${escapeHtml(zone)}</strong></div>
				<div role="listitem"><span>Projet VBA</span><strong>${probe.workbook.hasVbaProject ? 'Présent' : 'Absent'}</strong></div>
				<div role="listitem"><span>Signature VBA</span><strong>${escapeHtml(
					probe.workbook.vbaSignatureStatus === 'present'
						? 'Détectée'
						: probe.workbook.vbaSignatureStatus === 'absent'
							? 'Non détectée'
							: 'À confirmer'
				)}</strong></div>
				<div role="listitem"><span>Signature package</span><strong>${escapeHtml(
					probe.workbook.packageSignatureStatus === 'present'
						? 'Détectée'
						: probe.workbook.packageSignatureStatus === 'absent'
							? 'Non détectée'
							: 'À confirmer'
				)}</strong></div>
				<div role="listitem"><span>Chiffrement Office</span><strong>${probe.workbook.officePackageEncrypted ? 'Actif' : 'Non détecté'}</strong></div>
				<div role="listitem"><span>Chiffrement EFS</span><strong>${probe.workbook.efsEncrypted ? 'Actif' : 'Non détecté'}</strong></div>
				<div role="listitem"><span>Attribut lecture seule</span><strong>${probe.workbook.readOnly ? 'Actif' : 'Non'}</strong></div>
				<div role="listitem"><span>SHA-256 inspecté</span><strong><code>${escapeHtml(probe.workbook.sha256)}</code></strong></div>
			</div>
		</section>

		<section class="notice" aria-label="Limites du diagnostic">
			<strong>Diagnostic local en lecture seule.</strong>
			L’extension n’exécute aucune macro et ne modifie ni le registre, ni le Centre de gestion Microsoft 365, ni une stratégie d’entreprise.
		</section>

		<div class="toolbar" role="toolbar" aria-label="Actions du Centre de sécurité">
			<button class="primary" data-action="refresh">Actualiser</button>
			<button data-action="copyReport">Copier le rapport</button>
			<button data-action="openExcelSecurity">Ouvrir le classeur dans Excel</button>
			<button data-action="openExtensionSettings">Réglages de l’extension</button>
			<button data-action="openEnterpriseAdmin">Portail entreprise Microsoft 365 (admin)</button>
			<button data-action="openAdminDocs">Documentation administrateur</button>
		</div>
		<p class="admin-note">Le portail nécessite un compte disposant d’un rôle administrateur autorisé. L’extension n’obtient aucun droit supplémentaire et ne modifie aucune règle.</p>

		<section>
			<div class="section-heading">
				<div><p class="eyebrow">Décisions</p><h2>Règles applicables au classeur</h2></div>
				<span>${managedCount} réglage(s) géré(s)</span>
			</div>
			<div class="rule-grid">${findingRows}</div>
		</section>

		<section>
			<p class="eyebrow">Périmètre réel</p>
			<h2>Capacités de l’extension</h2>
			<div class="table-wrap"><table>
				<thead><tr><th>Fonction</th><th>État</th><th>Explication</th></tr></thead>
				<tbody>${capabilityRows}</tbody>
			</table></div>
		</section>

		<section>
			<div class="section-heading">
				<div><p class="eyebrow">Configuration Office ${escapeHtml(
					probe.office.version
				)}</p><h2>Niveaux détectés sur ce PC</h2></div>
				<span>${
					probe.office.cloudPolicyDetected
						? 'Règle Cloud Policy détectée'
						: probe.office.cloudPolicyServiceDetected
							? 'Service Cloud Policy détecté, aucune règle pertinente trouvée'
							: 'Aucune Cloud Policy détectée'
				}</span>
			</div>
			${
				settingRows
					? `<div class="table-wrap"><table>
						<thead><tr><th>Contrôle</th><th>Valeur</th><th>Source</th><th>Preuve locale</th></tr></thead>
						<tbody>${settingRows}</tbody>
					</table></div>`
					: '<p class="empty">Aucun réglage Office explicite n’a été trouvé dans les emplacements inspectés.</p>'
			}
		</section>

		<section>
			<p class="eyebrow">Confiance</p>
			<h2>Emplacements approuvés</h2>
			<p class="section-intro">Un emplacement approuvé peut contourner plusieurs contrôles Office. Il doit rester rare et administré.</p>
			${
				trustedLocationRows
					? `<div class="table-wrap"><table>
						<thead><tr><th>Chemin</th><th>Sous-dossiers</th><th>Source</th></tr></thead>
						<tbody>${trustedLocationRows}</tbody>
					</table></div>`
					: '<p class="empty">Aucun emplacement approuvé explicite n’a été détecté.</p>'
			}
		</section>

		<section>
			<p class="eyebrow">Classification</p>
			<h2>Étiquettes Microsoft Purview</h2>
			${sensitivityLabels}
		</section>

		<footer>Inspection effectuée : ${escapeHtml(
			probe.inspectedAtUtc || 'date indisponible'
		)}. Ce résultat décrit les éléments observables localement ; Excel reste l’autorité au moment d’ouvrir le fichier.</footer>`;
}

export class SecurityCenterPanel implements vscode.Disposable {
	private panel: vscode.WebviewPanel | undefined;
	private workbookUri: vscode.Uri | undefined;
	private report: EnterpriseSecurityReport | undefined;
	private inspectionSequence = 0;
	private readonly panelDisposables: vscode.Disposable[] = [];
	private readonly extensionSettingsQuery: string;

	constructor(
		context: vscode.ExtensionContext,
		private readonly service: OfficeSecurityService,
		private readonly outputChannel: vscode.OutputChannel,
		private readonly openExcelCallback: OpenExcelCallback
	) {
		const packageJson = context.extension.packageJSON as {
			publisher?: unknown;
			name?: unknown;
		};
		const publisher =
			typeof packageJson.publisher === 'string'
				? packageJson.publisher
				: 'steph-tools';
		const name =
			typeof packageJson.name === 'string'
				? packageJson.name
				: 'excel-ai-vba-studio';
		this.extensionSettingsQuery = `@ext:${publisher}.${name}`;
	}

	dispose(): void {
		this.inspectionSequence += 1;
		const panel = this.panel;
		this.panel = undefined;
		this.workbookUri = undefined;
		this.report = undefined;
		panel?.dispose();
		this.disposePanelListeners();
	}

	async open(workbookUri: vscode.Uri): Promise<void> {
		this.workbookUri = workbookUri;
		this.ensurePanel();
		if (!this.panel) {
			return;
		}
		this.panel.title = `Sécurité · ${path.basename(workbookUri.fsPath)}`;
		this.panel.reveal(vscode.ViewColumn.Active, false);
		await this.refresh();
	}

	private ensurePanel(): void {
		if (this.panel) {
			return;
		}
		const panel = vscode.window.createWebviewPanel(
			'excelAiVbaStudio.securityCenter',
			'Centre de sécurité Excel',
			vscode.ViewColumn.Active,
			{
				enableScripts: true,
				localResourceRoots: []
			}
		);
		this.panel = panel;
		this.panelDisposables.push(
			panel.onDidDispose(() => {
				if (this.panel === panel) {
					this.inspectionSequence += 1;
					this.panel = undefined;
					this.workbookUri = undefined;
					this.report = undefined;
				}
				this.disposePanelListeners();
			}),
			panel.webview.onDidReceiveMessage(message => {
				void this.handleMessage(message as SecurityCenterMessage).catch(error => {
					const text = error instanceof Error ? error.message : String(error);
					this.outputChannel.appendLine(
						`[security] Action du panneau impossible : ${text}`
					);
					void vscode.window.showErrorMessage(
						`Centre de sécurité Excel : ${text}`
					);
				});
			})
		);
	}

	private disposePanelListeners(): void {
		for (const disposable of this.panelDisposables.splice(0)) {
			disposable.dispose();
		}
	}

	private async refresh(): Promise<void> {
		const panel = this.panel;
		const workbookUri = this.workbookUri;
		if (!panel || !workbookUri) {
			return;
		}
		const sequence = ++this.inspectionSequence;
		panel.webview.html = this.getHtml(
			panel.webview,
			'<section class="loading" role="status"><div class="spinner"></div><h1>Analyse de la sécurité Office</h1><p>Lecture du fichier, de son origine Windows et des stratégies Office visibles…</p></section>'
		);
		try {
			const report = await this.service.inspect(workbookUri);
			if (
				sequence !== this.inspectionSequence ||
				this.panel !== panel ||
				this.workbookUri?.toString() !== workbookUri.toString()
			) {
				return;
			}
			this.report = report;
			panel.webview.html = this.getHtml(panel.webview, renderReport(report));
		} catch (error) {
			if (sequence !== this.inspectionSequence || this.panel !== panel) {
				return;
			}
			this.report = undefined;
			const message = error instanceof Error ? error.message : String(error);
			this.outputChannel.appendLine(`[security] Erreur : ${message}`);
			panel.webview.html = this.getHtml(
				panel.webview,
				`<section class="error-card"><p class="eyebrow">Diagnostic interrompu</p><h1>Le niveau de sécurité n’a pas pu être lu</h1><p>${escapeHtml(
					message
				)}</p><button class="primary" data-action="refresh">Réessayer</button></section>`
			);
			void vscode.window.showErrorMessage(
				`Centre de sécurité Excel : ${message}`
			);
		}
	}

	private async handleMessage(message: SecurityCenterMessage): Promise<void> {
		if (!message || typeof message.type !== 'string') {
			return;
		}
		switch (message.type) {
			case 'refresh':
				await this.refresh();
				return;
			case 'copyReport':
				if (!this.report) {
					void vscode.window.showWarningMessage(
						'Aucun rapport de sécurité n’est disponible.'
					);
					return;
				}
				await vscode.env.clipboard.writeText(
					formatEnterpriseSecurityReport(this.report)
				);
				void vscode.window.showInformationMessage(
					'Le rapport de sécurité a été copié.'
				);
				return;
			case 'openExcelSecurity':
				if (!this.workbookUri) {
					return;
				}
				try {
					const opened = await this.openExcelCallback(this.workbookUri);
					if (opened) {
						void vscode.window.showInformationMessage(
							'Dans Excel : Développeur > Sécurité des macros. Les options gérées par l’organisation restent verrouillées.'
						);
					}
				} catch (error) {
					const message = error instanceof Error ? error.message : String(error);
					this.outputChannel.appendLine(
						`[security] Ouverture d’Excel impossible : ${message}`
					);
					void vscode.window.showErrorMessage(
						`Ouverture d’Excel impossible : ${message}`
					);
				}
				return;
			case 'openExtensionSettings':
				await vscode.commands.executeCommand(
					'workbench.action.openSettings',
					this.extensionSettingsQuery
				);
				return;
			case 'openEnterpriseAdmin':
				if (!(await vscode.env.openExternal(MICROSOFT_365_APPS_ADMIN_URL))) {
					throw new Error('Le portail Microsoft 365 Apps n’a pas pu être ouvert.');
				}
				return;
			case 'openAdminDocs':
				if (!(await vscode.env.openExternal(MICROSOFT_CLOUD_POLICY_DOCS_URL))) {
					throw new Error('La documentation Microsoft n’a pas pu être ouverte.');
				}
				return;
			default:
				return;
		}
	}

	private getHtml(webview: vscode.Webview, content: string): string {
		const nonce = randomBytes(24).toString('base64');
		return `<!doctype html>
<html lang="fr">
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width,initial-scale=1">
	<meta http-equiv="Content-Security-Policy" content="default-src 'none'; connect-src 'none'; img-src 'none'; font-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}'; base-uri 'none'; form-action 'none';">
	<title>Centre de sécurité Excel</title>
	<style nonce="${nonce}">
		:root{color-scheme:light dark}
		*{box-sizing:border-box}
		body{margin:0;background:var(--vscode-editor-background);color:var(--vscode-editor-foreground);font-family:var(--vscode-font-family);font-size:var(--vscode-font-size);line-height:1.55}
		main{width:min(1180px,calc(100% - 40px));margin:0 auto;padding:28px 0 54px}
		h1,h2,h3,p{margin-top:0} h1{font-size:30px;line-height:1.15;margin-bottom:8px} h2{font-size:20px;margin-bottom:12px} h3{font-size:15px;margin-bottom:0}
		section{margin:0 0 28px}.eyebrow{margin:0 0 4px;text-transform:uppercase;letter-spacing:.09em;font-size:11px;font-weight:700;color:var(--vscode-descriptionForeground)}
		.hero,.workbook-card,.notice,.rule-card,.error-card{border:1px solid var(--vscode-panel-border);border-radius:10px;background:var(--vscode-sideBar-background)}
		.hero{display:flex;align-items:center;justify-content:space-between;padding:24px;border-left-width:6px}.hero.level-restricted{border-left-color:var(--vscode-testing-iconFailed,#d13438)}.hero.level-managed{border-left-color:var(--vscode-editorInfo-foreground,#0078d4)}.hero.level-standard{border-left-color:var(--vscode-testing-iconPassed,#16825d)}.hero.level-unknown{border-left-color:var(--vscode-editorWarning-foreground,#bf8803)}
		.summary{color:var(--vscode-descriptionForeground);margin:0}.level-mark{display:grid;place-items:center;flex:0 0 52px;height:52px;margin-left:18px;border:2px solid currentColor;border-radius:50%;font-size:24px;font-weight:800}
		.workbook-card{padding:20px}.path{font-family:var(--vscode-editor-font-family);word-break:break-all;color:var(--vscode-descriptionForeground)}
		.facts{display:grid;grid-template-columns:repeat(4,minmax(130px,1fr));gap:1px;margin-top:18px;background:var(--vscode-panel-border);border:1px solid var(--vscode-panel-border);border-radius:7px;overflow:hidden}.facts div{display:flex;flex-direction:column;padding:12px;background:var(--vscode-editor-background)}.facts span{font-size:11px;color:var(--vscode-descriptionForeground)}
		.notice{padding:14px 18px;border-color:var(--vscode-editorInfo-foreground);background:var(--vscode-textBlockQuote-background)}
		.toolbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:32px;position:sticky;top:0;z-index:2;padding:10px 0;background:var(--vscode-editor-background)}
		button{appearance:none;border:1px solid var(--vscode-button-border,transparent);border-radius:4px;padding:7px 12px;background:var(--vscode-button-secondaryBackground);color:var(--vscode-button-secondaryForeground);font:inherit;cursor:pointer}button:hover{background:var(--vscode-button-secondaryHoverBackground)}button:focus-visible{outline:2px solid var(--vscode-focusBorder);outline-offset:2px}button.primary{background:var(--vscode-button-background);color:var(--vscode-button-foreground)}button.primary:hover{background:var(--vscode-button-hoverBackground)}
		.section-heading,.rule-heading,.source-line{display:flex;align-items:center;justify-content:space-between;gap:12px}.section-heading>span,.source-line{color:var(--vscode-descriptionForeground);font-size:12px}
		.rule-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.rule-card{padding:16px}.rule-card p{margin:10px 0}.rule-card .impact{font-size:12px;color:var(--vscode-descriptionForeground)}
		.status,.managed-badge{display:inline-flex;align-items:center;width:max-content;border-radius:999px;padding:2px 8px;font-size:11px;font-weight:700;white-space:nowrap}.status-protected,.status-allowed{color:var(--vscode-testing-iconPassed,#16825d);background:color-mix(in srgb,var(--vscode-testing-iconPassed,#16825d) 14%,transparent)}.status-blocked{color:var(--vscode-testing-iconFailed,#d13438);background:color-mix(in srgb,var(--vscode-testing-iconFailed,#d13438) 14%,transparent)}.status-managed{color:var(--vscode-editorInfo-foreground,#0078d4);background:color-mix(in srgb,var(--vscode-editorInfo-foreground,#0078d4) 14%,transparent)}.status-prompt,.status-warning,.status-unknown{color:var(--vscode-editorWarning-foreground,#bf8803);background:color-mix(in srgb,var(--vscode-editorWarning-foreground,#bf8803) 14%,transparent)}.status-notApplicable{color:var(--vscode-descriptionForeground);background:var(--vscode-badge-background)}.managed-badge{color:var(--vscode-badge-foreground);background:var(--vscode-badge-background)}
		.table-wrap{overflow:auto;border:1px solid var(--vscode-panel-border);border-radius:8px}table{width:100%;border-collapse:collapse;min-width:680px}th,td{padding:10px 12px;border-bottom:1px solid var(--vscode-panel-border);text-align:left;vertical-align:top}th{position:sticky;top:0;background:var(--vscode-sideBarSectionHeader-background);font-size:12px}tr:last-child td{border-bottom:0}code{font-family:var(--vscode-editor-font-family);font-size:12px;word-break:break-all}details summary{cursor:pointer;color:var(--vscode-textLink-foreground)}
		.section-intro,.admin-note,.empty,footer{color:var(--vscode-descriptionForeground)}.admin-note{margin:-22px 0 32px;font-size:12px}.empty{padding:16px;border:1px dashed var(--vscode-panel-border);border-radius:8px}footer{border-top:1px solid var(--vscode-panel-border);padding-top:18px;font-size:12px}
		.loading,.error-card{max-width:620px;margin:15vh auto;padding:28px;text-align:center}.spinner{width:34px;height:34px;margin:0 auto 18px;border:3px solid var(--vscode-panel-border);border-top-color:var(--vscode-progressBar-background);border-radius:50%;animation:spin .8s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}
		@media (max-width:760px){main{width:min(calc(100% - 24px),1180px)}.facts,.rule-grid{grid-template-columns:1fr}.section-heading,.source-line{align-items:flex-start;flex-direction:column}.hero{align-items:flex-start}.level-mark{flex-basis:42px;height:42px;font-size:19px}}
		@media (prefers-reduced-motion:reduce){.spinner{animation:none}}
	</style>
</head>
<body>
	<main>${content}</main>
	<script nonce="${nonce}">
		(() => {
			const vscode = acquireVsCodeApi();
			const allowedActions = new Set([
				'refresh',
				'copyReport',
				'openExcelSecurity',
				'openExtensionSettings',
				'openEnterpriseAdmin',
				'openAdminDocs'
			]);
			document.addEventListener('click', event => {
				const target = event.target instanceof Element
					? event.target.closest('[data-action]')
					: null;
				const action = target instanceof HTMLElement ? target.dataset.action : undefined;
				if (action && allowedActions.has(action)) {
					vscode.postMessage({ type: action });
				}
			});
		})();
	</script>
</body>
</html>`;
	}
}
