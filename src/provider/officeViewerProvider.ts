import { getReactWebviewResourceRoots } from '@/common/extensionResource';
import { getFileSuffix } from '@/common/fileSuffix';
import { Handler } from '@/common/handler';
import { ReactApp } from '@/common/reactApp';
import { createHash } from 'crypto';
import * as vscode from 'vscode';
import {
	handleCommonEvent,
	type OfficeDocumentLifecycle,
	type OfficeEditorSession,
} from './compress/commonHandler';

const SUPPORTED_SPREADSHEET_SUFFIXES = new Set([
	'.xlsx',
	'.xlsm',
	'.xls',
	'.csv',
	'.tsv',
]);
const MAX_BACKUP_BYTES = 128 * 1024 * 1024;

interface SpreadsheetBackupEnvelope {
	version: 1;
	documentUri: string;
	sheets: unknown[];
	sourceSha256: string;
}

class OfficeCustomDocument
implements vscode.CustomDocument, OfficeDocumentLifecycle {
	private readonly sessions = new Set<OfficeEditorSession>();

	constructor(
		readonly uri: vscode.Uri,
		public restoredSheets: unknown[] | undefined,
		public restoredSourceSha256: string | undefined,
		private readonly fireChange: (
			event: vscode.CustomDocumentContentChangeEvent<OfficeCustomDocument>
		) => void
	) {}

	dispose(): void {
		this.sessions.clear();
	}

	markChanged(): void {
		this.fireChange({ document: this });
	}

	registerSession(session: OfficeEditorSession): vscode.Disposable {
		this.sessions.add(session);
		return new vscode.Disposable(() => this.sessions.delete(session));
	}

	hasSession(): boolean {
		return this.sessions.size > 0;
	}

	getSession(): OfficeEditorSession {
		const session = Array.from(this.sessions).at(-1);
		if (!session) {
			throw new Error('Spreadsheet editor is not ready.');
		}
		return session;
	}

	getSessions(): OfficeEditorSession[] {
		return Array.from(this.sessions);
	}
}

/**
 * Hosts the Excel/CSV webview. Other legacy Office viewers are intentionally
 * excluded from the published runtime.
 */
export class OfficeViewerProvider
implements vscode.CustomEditorProvider<OfficeCustomDocument> {
	private readonly changeEmitter = new vscode.EventEmitter<
		vscode.CustomDocumentContentChangeEvent<OfficeCustomDocument>
	>();
	readonly onDidChangeCustomDocument = this.changeEmitter.event;

	constructor(private readonly context: vscode.ExtensionContext) {}

	bindCustomEditor(
		viewOption: { webviewOptions: vscode.WebviewPanelOptions }
	): vscode.Disposable {
		const registration = vscode.window.registerCustomEditorProvider(
			'excelAiVbaStudio.officeViewer',
			this,
			viewOption
		);
		return vscode.Disposable.from(registration, this.changeEmitter);
	}

	async openCustomDocument(
		uri: vscode.Uri,
		openContext: vscode.CustomDocumentOpenContext,
		token: vscode.CancellationToken
	): Promise<OfficeCustomDocument> {
		let restoredSheets: unknown[] | undefined;
		let restoredSourceSha256: string | undefined;
		if (openContext.backupId) {
			const backupUri = vscode.Uri.parse(openContext.backupId, true);
			const stat = await vscode.workspace.fs.stat(backupUri);
			if (stat.size <= 0 || stat.size > MAX_BACKUP_BYTES) {
				throw new Error('Spreadsheet recovery backup has an invalid size.');
			}
			const bytes = await vscode.workspace.fs.readFile(backupUri);
			if (token.isCancellationRequested) {
				throw new vscode.CancellationError();
			}
			const parsed = JSON.parse(
				new TextDecoder('utf-8', { fatal: true }).decode(bytes)
			) as Partial<SpreadsheetBackupEnvelope>;
			if (
				parsed.version !== 1 ||
				parsed.documentUri !== uri.toString() ||
				!Array.isArray(parsed.sheets) ||
				!/^[0-9a-f]{64}$/.test(parsed.sourceSha256 ?? '')
			) {
				throw new Error('Spreadsheet recovery backup is invalid.');
			}
			const currentBytes = await vscode.workspace.fs.readFile(uri);
			if (token.isCancellationRequested) {
				throw new vscode.CancellationError();
			}
			const currentSha256 = createHash('sha256')
				.update(currentBytes)
				.digest('hex');
			if (currentSha256 !== parsed.sourceSha256) {
				throw new Error(
					'Spreadsheet recovery was stopped because the source file ' +
					'changed after this backup was created. The recovery backup ' +
					'was retained to prevent overwriting newer data.'
				);
			}
			restoredSheets = parsed.sheets;
			restoredSourceSha256 = parsed.sourceSha256;
		}
		return new OfficeCustomDocument(
			uri,
			restoredSheets,
			restoredSourceSha256,
			event => this.changeEmitter.fire(event)
		);
	}

	resolveCustomEditor(
		document: OfficeCustomDocument,
		webviewPanel: vscode.WebviewPanel
	): void | Thenable<void> {
		const uri = document.uri;
		const suffix = getFileSuffix(uri.fsPath);
		if (!SUPPORTED_SPREADSHEET_SUFFIXES.has(suffix)) {
			void vscode.commands.executeCommand('vscode.openWith', uri, 'default');
			return;
		}
		if (document.hasSession()) {
			throw new Error(
				'Only one editable view can be open for a spreadsheet. ' +
				'Close the existing split view before opening another.'
			);
		}

		const webview = webviewPanel.webview;
		const folderPath = vscode.Uri.joinPath(uri, '..');
		webview.options = {
			enableScripts: true,
			localResourceRoots: [
				...getReactWebviewResourceRoots(this.context),
				folderPath,
			],
		};

		const handler = Handler.bind(webviewPanel, uri);
		handleCommonEvent(uri, handler, document);
		return ReactApp.view(webview, { route: 'excel' });
	}

	async saveCustomDocument(
		document: OfficeCustomDocument,
		cancellation: vscode.CancellationToken
	): Promise<void> {
		if (cancellation.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		await document.getSession().save();
		document.restoredSheets = undefined;
		document.restoredSourceSha256 = undefined;
	}

	async saveCustomDocumentAs(
		document: OfficeCustomDocument,
		destination: vscode.Uri,
		cancellation: vscode.CancellationToken
	): Promise<void> {
		if (cancellation.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		const suffix = getFileSuffix(destination.fsPath);
		if (suffix === '.xlsm' || suffix === '.xls') {
			throw new Error(
				'Save As is blocked for macro-enabled and legacy workbooks. ' +
				'Use native Excel so VBA and opaque workbook data remain intact.'
			);
		}
		await document.getSession().saveAs(destination);
		document.restoredSheets = undefined;
		document.restoredSourceSha256 = undefined;
	}

	async revertCustomDocument(
		document: OfficeCustomDocument,
		cancellation: vscode.CancellationToken
	): Promise<void> {
		if (cancellation.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		await Promise.all(
			document.getSessions().map(session => session.revert())
		);
		document.restoredSheets = undefined;
		document.restoredSourceSha256 = undefined;
	}

	async backupCustomDocument(
		document: OfficeCustomDocument,
		backupContext: vscode.CustomDocumentBackupContext,
		cancellation: vscode.CancellationToken
	): Promise<vscode.CustomDocumentBackup> {
		const backup = await document.getSession().backup();
		if (cancellation.isCancellationRequested) {
			throw new vscode.CancellationError();
		}
		if (
			!Array.isArray(backup.sheets) ||
			!/^[0-9a-f]{64}$/.test(backup.sourceSha256)
		) {
			throw new Error('Spreadsheet editor returned an invalid backup state.');
		}
		const envelope: SpreadsheetBackupEnvelope = {
			version: 1,
			documentUri: document.uri.toString(),
			sheets: backup.sheets,
			sourceSha256: backup.sourceSha256,
		};
		const bytes = new TextEncoder().encode(JSON.stringify(envelope));
		if (bytes.byteLength <= 0 || bytes.byteLength > MAX_BACKUP_BYTES) {
			throw new Error('Spreadsheet recovery backup exceeds 128 MiB.');
		}
		await vscode.workspace.fs.createDirectory(
			vscode.Uri.joinPath(backupContext.destination, '..')
		);
		await vscode.workspace.fs.writeFile(backupContext.destination, bytes);
		return {
			id: backupContext.destination.toString(),
			delete: () => {
				void vscode.workspace.fs
					.delete(backupContext.destination)
					.then(undefined, () => undefined);
			},
		};
	}
}
