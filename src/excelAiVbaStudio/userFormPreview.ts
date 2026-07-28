import { randomBytes } from 'crypto';
import * as path from 'path';
import * as vscode from 'vscode';

interface UserFormNode {
	type: string;
	name: string;
	properties: Record<string, string>;
	children: UserFormNode[];
}

const MAX_CONTROLS = 200;

function escapeHtml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

function decodeVbaString(value: string): string {
	const trimmed = value.trim();
	const quoted = /^"(.*)"(?:\s*&\s*ChrW\(\d+\))?$/i.exec(trimmed);
	return quoted ? quoted[1].replace(/""/g, '"') : trimmed;
}

function parseUserForm(source: string): UserFormNode | undefined {
	const stack: UserFormNode[] = [];
	let root: UserFormNode | undefined;
	let controlCount = 0;
	let ignoredDepth = 0;

	for (const rawLine of source.split(/\r?\n/)) {
		const line = rawLine.trim();
		const begin = /^Begin\s+([\w.]+)\s+([^\s]+)$/i.exec(line);
		if (begin) {
			if (ignoredDepth > 0 || controlCount >= MAX_CONTROLS) {
				ignoredDepth++;
				continue;
			}
			const node: UserFormNode = {
				type: begin[1],
				name: begin[2],
				properties: {},
				children: []
			};
			const parent = stack[stack.length - 1];
			if (parent) {
				parent.children.push(node);
			} else if (/userform$/i.test(node.type)) {
				root = node;
			}
			stack.push(node);
			controlCount++;
			continue;
		}
		if (/^End$/i.test(line)) {
			if (ignoredDepth > 0) {
				ignoredDepth--;
			} else {
				stack.pop();
			}
			continue;
		}
		const current = stack[stack.length - 1];
		const property = /^([A-Za-z][\w.]*)\s*=\s*(.+)$/.exec(line);
		if (current && property) {
			current.properties[property[1]] = decodeVbaString(property[2]);
		}
	}
	return root;
}

function numericProperty(
	node: UserFormNode,
	keys: string[],
	fallback: number
): number {
	for (const key of keys) {
		const parsed = Number(node.properties[key]);
		if (Number.isFinite(parsed)) {
			return parsed;
		}
	}
	return fallback;
}

function controlLabel(node: UserFormNode): string {
	return node.properties.Caption || node.properties.Text || node.properties.Value || node.name;
}

function controlClass(node: UserFormNode): string {
	const type = node.type.toLocaleLowerCase('en-US');
	if (type.includes('commandbutton')) {
		return 'button';
	}
	if (type.includes('togglebutton')) {
		return 'toggle';
	}
	if (type.includes('textbox')) {
		return 'input';
	}
	if (type.includes('combobox')) {
		return 'combo';
	}
	if (type.includes('checkbox')) {
		return 'checkbox';
	}
	if (type.includes('optionbutton')) {
		return 'radio';
	}
	if (type.includes('listbox')) {
		return 'list';
	}
	if (type.includes('image')) {
		return 'image';
	}
	if (type.includes('spinbutton')) {
		return 'spin';
	}
	if (type.includes('scrollbar')) {
		return 'scrollbar';
	}
	if (type.includes('frame')) {
		return 'frame';
	}
	return 'label';
}

function booleanProperty(
	node: UserFormNode,
	key: string,
	fallback: boolean
): boolean {
	const raw = node.properties[key];
	if (raw === undefined) {
		return fallback;
	}
	return !/^(?:0|false)$/i.test(raw.trim());
}

function renderControl(node: UserFormNode): string {
	const left = numericProperty(node, ['Left'], 12) / 15;
	const top = numericProperty(node, ['Top'], 12) / 15;
	const width = Math.max(24, numericProperty(node, ['Width'], 1200) / 15);
	const height = Math.max(18, numericProperty(node, ['Height'], 300) / 15);
	const label = escapeHtml(controlLabel(node));
	const type = escapeHtml(node.type.replace(/^.*\./, ''));
	const name = escapeHtml(node.name);
	const controlKind = controlClass(node);
	const children = node.children.map(renderControl).join('');
	const disabled = booleanProperty(node, 'Enabled', true) ? '' : ' disabled';
	const hidden = booleanProperty(node, 'Visible', true) ? '' : ' hidden';
	const attributes = `class="control ${controlKind}" data-control-name="${name}" data-control-type="${type}" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px" title="${type}: ${name}"${hidden}`;

	switch (controlKind) {
		case 'button':
			return `<button type="button" ${attributes}${disabled}>${label}</button>`;
		case 'toggle':
			return `<button type="button" ${attributes}${disabled} aria-pressed="false">${label}</button>`;
		case 'input':
			return `<input type="text" ${attributes}${disabled} value="${label}" aria-label="${name}">`;
		case 'combo':
			return `<select ${attributes}${disabled} aria-label="${name}"><option>${label}</option></select>`;
		case 'list':
			return `<select ${attributes}${disabled} size="2" aria-label="${name}"><option>${label}</option></select>`;
		case 'checkbox':
			return `<label ${attributes}><input type="checkbox"${disabled}> <span>${label}</span></label>`;
		case 'radio':
			return `<label ${attributes}><input type="radio" name="${name}"${disabled}> <span>${label}</span></label>`;
		case 'spin':
			return `<input type="number" ${attributes}${disabled} value="0" aria-label="${name}">`;
		case 'scrollbar':
			return `<input type="range" ${attributes}${disabled} min="0" max="100" value="0" aria-label="${name}">`;
		case 'frame':
			return `<fieldset ${attributes}${disabled}><legend>${label}</legend>${children}</fieldset>`;
		case 'image':
			return `<div ${attributes} role="img" aria-label="${label}">${label}</div>`;
		default:
			return `<div ${attributes}><span>${label}</span>${children}</div>`;
	}
}

function previewHtml(sourcePath: string, source: string, nonce: string): string {
	const form = parseUserForm(source);
	const title = form?.properties.Caption || form?.name || path.basename(sourcePath);
	const content = form
		? (() => {
				const width = Math.min(
					960,
					Math.max(
						320,
						numericProperty(form, ['ClientWidth', 'Width'], 6000) / 15
					)
				);
				const height = Math.min(
					720,
					Math.max(
						220,
						numericProperty(form, ['ClientHeight', 'Height'], 4200) / 15
					)
				);
				const controls = form.children.map(renderControl).join('');
				return `<section class="form" style="width:${width}px;height:${height}px">
<div class="title">${escapeHtml(title)}</div>
<div class="surface">${controls || '<div class="notice">Aucun contrôle visuel détecté.</div>'}</div>
</section>`;
		  })()
		: `<section class="empty">
<h2>${escapeHtml(path.basename(sourcePath))}</h2>
<p>Le fichier contient le code du UserForm, mais pas encore ses métadonnées visuelles. Le code reste ouvert dans l’éditeur VS Code et accessible à Copilot.</p>
</section>`;

	return `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'nonce-${nonce}';">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:var(--vscode-font-family);color:var(--vscode-foreground);background:var(--vscode-editor-background)}
header{min-height:44px;display:flex;gap:12px;align-items:center;padding:6px 14px;border-bottom:1px solid var(--vscode-panel-border)}
header span{opacity:.7}
header .spacer{flex:1}
header button{color:var(--vscode-button-foreground);background:var(--vscode-button-background);border:0;padding:6px 10px;cursor:pointer}
header button:hover{background:var(--vscode-button-hoverBackground)}
.interaction-status{margin:0;padding:8px 14px;border-bottom:1px solid var(--vscode-panel-border);color:var(--vscode-descriptionForeground);background:var(--vscode-textBlockQuote-background)}
main{padding:20px;overflow:auto}
.form{position:relative;min-width:320px;background:#f0f0f0;border:1px solid #777;box-shadow:0 10px 35px #0004;color:#111}
.title{height:28px;padding:5px 9px;color:#fff;background:#1769aa;font:12px "Segoe UI",sans-serif}
.surface{position:absolute;left:0;right:0;top:28px;bottom:0}
.control{position:absolute;border:1px solid transparent;padding:2px 4px;overflow:hidden;font:12px "Segoe UI",sans-serif;color:#111}
.control[hidden]{display:none}
.control.button,.control.toggle,.control.input,.control.combo,.control.list,.control.spin,.control.scrollbar{background:#fafafa;border-color:#888}
.control.button,.control.toggle{cursor:pointer}
.control.button:active,.control.toggle[aria-pressed="true"]{background:#d5e8f8;box-shadow:inset 0 1px 3px #0005}
.control.input{justify-content:flex-start;background:white}
.control.checkbox,.control.radio{display:flex;align-items:center;gap:4px}
.control.frame{border-color:#999;background:transparent;padding:10px 4px 4px}
.control.frame legend{padding:0 4px}
.control.image{border:1px dashed #888;background:#ddd}
.notice{padding:24px;color:#555}
.empty{max-width:680px}
</style>
</head>
<body>
<header>
<strong>Aperçu UserForm</strong>
<span>${escapeHtml(path.basename(sourcePath))}</span>
<span class="spacer"></span>
<button type="button" id="open-excel">Ouvrir dans Excel</button>
<button type="button" id="open-vbe">Ouvrir le VBE</button>
</header>
<p class="interaction-status" id="interaction-status">Aperçu interactif : les contrôles réagissent ici, mais aucun code VBA n’est exécuté. Utilisez Excel pour tester les événements réels.</p>
<main>${content}</main>
<script nonce="${nonce}">
const vscode = acquireVsCodeApi();
const status = document.getElementById('interaction-status');
document.getElementById('open-excel').addEventListener('click', () => {
	vscode.postMessage({ type: 'openExcel' });
});
document.getElementById('open-vbe').addEventListener('click', () => {
	vscode.postMessage({ type: 'openVbe' });
});
document.addEventListener('click', event => {
	const target = event.target;
	const element = target instanceof Element
		? target.closest('[data-control-name]')
		: null;
	if (!element || element.disabled) {
		return;
	}
	if (element.classList.contains('toggle')) {
		const pressed = element.getAttribute('aria-pressed') === 'true';
		element.setAttribute('aria-pressed', pressed ? 'false' : 'true');
	}
	const name = element.dataset.controlName || 'sans nom';
	status.textContent = 'Contrôle ' + name + ' activé dans l’aperçu. Aucun code VBA exécuté.';
});
</script>
</body>
</html>`;
}

export async function showUserFormPreview(
	sourceUri: vscode.Uri,
	source: string,
	workbookUri: vscode.Uri
): Promise<void> {
	const panel = vscode.window.createWebviewPanel(
		'excelAiVbaStudio.userFormPreview',
		`UserForm · ${path.basename(sourceUri.fsPath, path.extname(sourceUri.fsPath))}`,
		vscode.ViewColumn.Beside,
		{
			enableScripts: true,
			localResourceRoots: [],
			retainContextWhenHidden: true
		}
	);
	const receiver = panel.webview.onDidReceiveMessage(
		async (message: unknown) => {
			if (!message || typeof message !== 'object') {
				return;
			}
			const type = (message as { type?: unknown }).type;
			if (type === 'openExcel') {
				await vscode.commands.executeCommand(
					'excelAiVbaStudio.openExcel',
					workbookUri
				);
			} else if (type === 'openVbe') {
				await vscode.commands.executeCommand(
					'excelAiVbaStudio.openVbe',
					workbookUri
				);
			}
		}
	);
	panel.onDidDispose(() => receiver.dispose());
	panel.webview.html = previewHtml(
		sourceUri.fsPath,
		source,
		randomBytes(16).toString('hex')
	);
}
