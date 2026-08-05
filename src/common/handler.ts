import * as vscode from 'vscode';
import { SimpleEventEmitter } from "./simpleEventEmitter";
import { WebviewPanel } from "vscode";
import { Output } from "./Output";
import { errorMessage } from './errorMessage';

interface WebviewMessage {
    type: string;
    content?: unknown;
}

function isWebviewMessage(value: unknown): value is WebviewMessage {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
        return false;
    }
    const type = (value as { type?: unknown }).type;
    return typeof type === 'string' && /^[A-Za-z][A-Za-z0-9._-]{0,79}$/.test(type);
}

export class Handler {

    private isEnd: boolean;
    constructor(public panel: WebviewPanel, private eventEmitter: SimpleEventEmitter) { }

    on(event: string, callback: (content: any) => any | Promise<any>): this {
        if (event != 'init') {
            const listens = this.eventEmitter.listeners(event)
            if (listens.length >= 1) {
                this.eventEmitter.removeListener(event, listens[0])
            }
        }
        this.eventEmitter.on(event, async (content: any) => {
            try {
                await callback(content)
            } catch (error) {
                Output.debug(error)
                void Promise.resolve()
                    .then(() => vscode.window.showErrorMessage(errorMessage(error)))
                    .catch(reportingError => {
                        Output.debug(`Unable to show webview error: ${errorMessage(reportingError)}`)
                    })
            }
        })
        return this;
    }

    emit(event: string, content?: any) {
        if (this.isEnd) return this;
        void Promise.resolve()
            .then(() => this.panel.webview.postMessage({ type: event, content }))
            .catch(error => {
                Output.debug(`Unable to post webview event ${event}: ${errorMessage(error)}`)
            })
        return this;
    }

    public static bind(panel: WebviewPanel, uri: vscode.Uri): Handler {
        const eventEmitter = new SimpleEventEmitter();

        const fileWatcher = Handler.createFileWatcher(uri);
        fileWatcher?.onDidChange(e => {
            eventEmitter.emit("fileChange", e)
        })

        const changeDocumentSubscription = vscode.workspace.onDidChangeTextDocument(e => {
            if (e.document.uri.toString() === uri.toString() && e.contentChanges.length > 0) {
                eventEmitter.emit("externalUpdate", e)
            }
        });
        const handle = new Handler(panel, eventEmitter)
        panel.onDidDispose(() => {
            handle.isEnd = true;
            fileWatcher?.dispose()
            changeDocumentSubscription.dispose()
            eventEmitter.emit("dispose")
        });

        // bind from webview
        panel.webview.onDidReceiveMessage((message: unknown) => {
            if (!isWebviewMessage(message) || eventEmitter.listeners(message.type).length === 0) {
                return;
            }
            eventEmitter.emit(message.type, message.content)
        })
        return handle;
    }

    private static createFileWatcher(uri: vscode.Uri): vscode.FileSystemWatcher | undefined {
        const folder = vscode.workspace.getWorkspaceFolder(uri);
        if (folder) {
            const relativePath = vscode.workspace.asRelativePath(uri, false);
            return vscode.workspace.createFileSystemWatcher(
                new vscode.RelativePattern(folder, relativePath)
            );
        }
        if (uri.scheme === 'file') {
            return vscode.workspace.createFileSystemWatcher(uri.fsPath);
        }
        return undefined;
    }

}
