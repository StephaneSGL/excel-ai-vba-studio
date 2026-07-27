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
	if (type.includes('commandbutton') || type.includes('togglebutton')) {
		return 'button';
	}
	if (type.includes('textbox') || type.includes('combobox')) {
		return 'input';
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
	if (type.includes('frame')) {
		return 'frame';
	}
	return 'label';
}

function renderControl(node: UserFormNode): string {
	const left = numericProperty(node, ['Left'], 12) / 15;
	const top = numericProperty(node, ['Top'], 12) / 15;
	const width = Math.max(24, numericProperty(node, ['Width'], 1200) / 15);
	const height = Math.max(18, numericProperty(node, ['Height'], 300) / 15);
	const label = escapeHtml(controlLabel(node));
	const type = escapeHtml(node.type.replace(/^.*\./, ''));
	const children = node.children.map(renderControl).join('');
	return `<div class="control ${controlClass(node)}" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px" title="${type}: ${escapeHtml(node.name)}"><span>${label}</span>${children}</div>`;
}

function previewHtml(sourcePath: string, source: string): string {
	const form = parseUserForm(source);
	const title = form?.properties.Caption || form?.name || path.basename(sourcePath);
	if (!form) {
		return `<!doctype html>
<html lang="fr"><meta charset="utf-8"><body class="empty">
<h2>${escapeHtml(path.basename(sourcePath))}</h2>
<p>Le fichier contient le code du UserForm, mais pas encore ses métadonnées visuelles. Le code reste ouvert dans l’éditeur VS Code et accessible à Copilot.</p>
</body><style>
body{font-family:var(--vscode-font-family);color:var(--vscode-foreground);background:var(--vscode-editor-background);padding:24px}
.empty{max-width:680px}
</style></html>`;
	}

	const width = Math.min(
		960,
		Math.max(320, numericProperty(form, ['ClientWidth', 'Width'], 6000) / 15)
	);
	const height = Math.min(
		720,
		Math.max(220, numericProperty(form, ['ClientHeight', 'Height'], 4200) / 15)
	);
	const controls = form.children.map(renderControl).join('');
	return `<!doctype html>
<html lang="fr"><meta charset="utf-8"><body>
<header><strong>Aperçu UserForm</strong><span>${escapeHtml(path.basename(sourcePath))}</span></header>
<main>
<section class="form" style="width:${width}px;height:${height}px">
<div class="title">${escapeHtml(title)}</div>
<div class="surface">${controls || '<div class="notice">Aucun contrôle visuel détecté.</div>'}</div>
</section>
</main>
</body><style>
*{box-sizing:border-box}
body{margin:0;font-family:var(--vscode-font-family);color:var(--vscode-foreground);background:var(--vscode-editor-background)}
header{height:40px;display:flex;gap:12px;align-items:center;padding:0 14px;border-bottom:1px solid var(--vscode-panel-border)}
header span{opacity:.7}
main{padding:20px;overflow:auto}
.form{position:relative;min-width:320px;background:#f0f0f0;border:1px solid #777;box-shadow:0 10px 35px #0004;color:#111}
.title{height:28px;padding:5px 9px;color:#fff;background:#1769aa;font:12px "Segoe UI",sans-serif}
.surface{position:absolute;left:0;right:0;top:28px;bottom:0}
.control{position:absolute;border:1px solid transparent;padding:2px 4px;overflow:hidden;font:12px "Segoe UI",sans-serif}
.control.button,.control.input,.control.list{display:flex;align-items:center;justify-content:center;background:#fafafa;border-color:#888}
.control.input{justify-content:flex-start;background:white}
.control.checkbox span::before{content:"☐ ";font-size:15px}
.control.radio span::before{content:"○ ";font-size:15px}
.control.frame{border-color:#999;background:transparent}
.control.image{border:1px dashed #888;background:#ddd}
.notice{padding:24px;color:#555}
</style></html>`;
}

export async function showUserFormPreview(
	sourceUri: vscode.Uri,
	source: string
): Promise<void> {
	const panel = vscode.window.createWebviewPanel(
		'excelAiVbaStudio.userFormPreview',
		`UserForm · ${path.basename(sourceUri.fsPath, path.extname(sourceUri.fsPath))}`,
		vscode.ViewColumn.Beside,
		{
			enableScripts: false,
			retainContextWhenHidden: true
		}
	);
	panel.webview.html = previewHtml(sourceUri.fsPath, source);
}
