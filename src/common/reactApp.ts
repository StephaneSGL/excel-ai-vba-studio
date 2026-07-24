import * as vscode from 'vscode';
import {
    encodeHtmlAttributeJson,
    extensionResource,
    getExtensionUri,
    readExtensionText,
    withWebviewCsp,
} from './extensionResource';

interface ViewOption {
    route: string;
}

export class ReactApp {

    private static context: vscode.ExtensionContext;
    private static webviewUri: vscode.Uri;
    public static IS_DEV = false;

    public static init(context: vscode.ExtensionContext) {
        this.context = context;
        this.webviewUri = extensionResource(context, 'out', 'webview');
        this.IS_DEV = context.extensionMode == vscode.ExtensionMode.Development;
    }

    public static async view(webview: vscode.Webview, option: ViewOption) {
        const html = await this.readContent();
        const devServerUrl = this.IS_DEV ? 'http://127.0.0.1:5739' : undefined;
        webview.html = withWebviewCsp(this.buildPath(html, webview), webview, {
            developmentServer: devServerUrl,
        })
            .replace(`{{configs}}`, encodeHtmlAttributeJson({
                ...option,
                language: vscode.env.language,
                config: vscode.workspace.getConfiguration('excelAiVbaStudio')
            }));
    }

    private static async readContent(): Promise<string> {
        if (this.IS_DEV) {
            const devServerUrl = 'http://127.0.0.1:5739';
            const response = await fetch(`${devServerUrl}/index.html`);
            if (!response.ok) {
                throw new Error(`Vite dev server returned HTTP ${response.status}.`);
            }
            const data = await response.text();
            return data.replace(/(["'])\/(?=(?:@|src\/|index\.html\?))/g, `$1${devServerUrl}/`);
        }
        return readExtensionText(this.context, 'out', 'webview', 'index.html');
    }

    private static buildPath(data: string, webview: vscode.Webview): string {
        const baseUrl = ReactApp.getBaseUrl(webview);
        return data.replace('<base href="/">', `<base href="${baseUrl}/">`);
    }

    private static getBaseUrl(webview: vscode.Webview) {
        if (this.IS_DEV) {
            return `http://127.0.0.1:5739`;
        }
        return webview.asWebviewUri(this.webviewUri).toString();
    }

    public static getExtensionUri(): vscode.Uri {
        return getExtensionUri(this.context);
    }

}
