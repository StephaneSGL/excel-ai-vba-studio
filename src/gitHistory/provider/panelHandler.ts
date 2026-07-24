import { SimpleEventEmitter } from '../../common/simpleEventEmitter';
import { WebviewPanel } from 'vscode';

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

export class PanelHandler {
    constructor(
        public readonly panel: WebviewPanel,
        private readonly eventEmitter: SimpleEventEmitter
    ) { }

    on(event: string, callback: (content: unknown) => void | Promise<void>): this {
        if (event !== 'ready') {
            const listeners = this.eventEmitter.listeners(event);
            if (listeners.length >= 1) {
                this.eventEmitter.removeListener(event, listeners[0] as (...args: unknown[]) => void);
            }
        }
        this.eventEmitter.on(event, async (content: unknown) => {
            await callback(content);
        });
        return this;
    }

    emit(event: string, content?: unknown): this {
        this.panel.webview.postMessage({ type: event, content });
        return this;
    }

    static bind(panel: WebviewPanel): PanelHandler {
        const eventEmitter = new SimpleEventEmitter();
        panel.onDidDispose(() => eventEmitter.emit('dispose'));
        panel.webview.onDidReceiveMessage((message: unknown) => {
            if (!isWebviewMessage(message) || eventEmitter.listeners(message.type).length === 0) {
                return;
            }
            eventEmitter.emit(message.type, message.content);
        });
        return new PanelHandler(panel, eventEmitter);
    }
}
