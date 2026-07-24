import * as vscode from 'vscode';

export const EXTENSION_HOST_WEB_CONTEXT = 'excelAiVbaStudio.extensionHost.web';

let webExtensionHost = false;

export function setExtensionHostContext(): void {
    webExtensionHost = true;
    void vscode.commands.executeCommand('setContext', EXTENSION_HOST_WEB_CONTEXT, true);
}

export function isWebExtensionHost(): boolean {
    return webExtensionHost;
}
