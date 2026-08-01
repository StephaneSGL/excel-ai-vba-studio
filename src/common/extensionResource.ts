import * as vscode from 'vscode';

export function getExtensionUri(context: vscode.ExtensionContext): vscode.Uri {
    return context.extensionUri ?? vscode.Uri.file(context.extensionPath);
}

export function extensionResource(context: vscode.ExtensionContext, ...segments: string[]): vscode.Uri {
    return vscode.Uri.joinPath(getExtensionUri(context), ...segments);
}

export async function readExtensionText(context: vscode.ExtensionContext, ...segments: string[]): Promise<string> {
    const bytes = await vscode.workspace.fs.readFile(extensionResource(context, ...segments));
    return new TextDecoder().decode(bytes);
}

export function getReactWebviewResourceRoots(context: vscode.ExtensionContext): vscode.Uri[] {
    return [
        extensionResource(context, 'out', 'webview'),
        extensionResource(context, 'icons'),
    ];
}

export interface WebviewCspOptions {
    allowRemoteImages?: boolean;
    allowDataFrames?: boolean;
    developmentServer?: string;
}

function escapeHtmlAttribute(value: string): string {
    return value
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

export function escapeHtml(value: string): string {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

export function encodeHtmlAttributeJson(value: unknown): string {
    return escapeHtmlAttribute(JSON.stringify(value));
}

export function withWebviewCsp(
    html: string,
    webview: vscode.Webview,
    options: WebviewCspOptions = {},
): string {
    const dev = options.developmentServer?.trim();
    const scriptSources = [webview.cspSource];
    const styleSources = [webview.cspSource, "'unsafe-inline'"];
    const connectSources = [webview.cspSource];
    if (dev) {
        scriptSources.push(dev);
        styleSources.push(dev);
        connectSources.push(dev, dev.replace(/^http/i, 'ws'));
    }
    const imageSources = [
        webview.cspSource,
        'data:',
        'blob:',
        ...(options.allowRemoteImages ? ['https:', 'http:'] : []),
    ];
    const frameSources = options.allowDataFrames ? ' frame-src data: blob:;' : '';
    const policy = [
        "default-src 'none'",
        `img-src ${imageSources.join(' ')}`,
        `font-src ${webview.cspSource} data:`,
        `style-src ${styleSources.join(' ')}`,
        `script-src ${scriptSources.join(' ')}`,
        `connect-src ${connectSources.join(' ')}`,
        `worker-src ${webview.cspSource} blob:`,
    ].join('; ') + `;${frameSources}`;
    const meta = `<meta http-equiv="Content-Security-Policy" content="${escapeHtmlAttribute(policy)}">`;
    return /<head[\s>]/i.test(html)
        ? html.replace(/<head(\s[^>]*)?>/i, (head) => `${head}\n    ${meta}`)
        : `${meta}\n${html}`;
}
