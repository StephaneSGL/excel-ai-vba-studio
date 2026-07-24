import * as vscode from 'vscode';
import { CommitService } from '../service/commitService';
import { GitActions } from '../service/gitActions';
import { findGit } from '../service/findGit';
import { GitExecutor } from '../service/gitExecutor';
import { GitRepoCommands } from '../service/gitRepoCommands';
import { RepoDiscovery } from '../service/repoDiscovery';
import { GitActionHandler } from './gitActionHandler';
import { GitHistoryPanel, GIT_HISTORY_VIEW_TYPE } from './gitHistoryPanel';
import type { GitHistoryPanelContext } from './gitHistoryPanelContext';
import { GitHistoryPanelSerializer } from './gitHistoryPanelSerializer';
import {
    buildPanelContextFromCommandArg,
    resolvePreferredRepo,
} from '../util/resolveGitHistoryCommandContext';
import { normalizeRepoPath } from '../util/repoPath';
import { i18n } from '@/common/global';

let commitService: CommitService | undefined;
let repoDiscovery: RepoDiscovery | undefined;
let gitActions: GitActions | undefined;
let gitRepoCommands: GitRepoCommands | undefined;
let gitActionHandler: GitActionHandler | undefined;

interface GitRepository {
    rootUri: vscode.Uri;
    inputBox?: { value: string };
}

interface GitApi {
    repositories: ReadonlyArray<GitRepository>;
}

interface GitExtensionExports {
    getAPI(version: 1): GitApi;
}

function resolveFileUri(arg?: vscode.Uri | { resourceUri?: vscode.Uri }): vscode.Uri | undefined {
    if (arg && typeof arg === 'object' && 'resourceUri' in arg && arg.resourceUri?.scheme === 'file') {
        return arg.resourceUri;
    }
    if (arg && typeof arg === 'object' && 'scheme' in arg && arg.scheme === 'file') {
        return arg as vscode.Uri;
    }
    const active = vscode.window.activeTextEditor?.document.uri;
    if (active?.scheme === 'file') {
        return active;
    }
    return undefined;
}

function mergePanelContext(
    base: GitHistoryPanelContext,
    fromArg: GitHistoryPanelContext,
): GitHistoryPanelContext {
    return {
        fileUri: base.fileUri ?? fromArg.fileUri,
        preferredRepo: base.preferredRepo ?? fromArg.preferredRepo,
    };
}

async function getGitInputCommitMessage(repo: string): Promise<string> {
    const extension = vscode.extensions.getExtension<GitExtensionExports>('vscode.git');
    if (!extension) {
        return '';
    }
    if (!extension.isActive) {
        try {
            await extension.activate();
        } catch {
            return '';
        }
    }
    const api = extension.exports?.getAPI(1);
    const repoPath = normalizeRepoPath(repo);
    const gitRepo = api?.repositories.find((candidate) =>
        candidate.rootUri.scheme === 'file' && normalizeRepoPath(candidate.rootUri.fsPath) === repoPath
    );
    return gitRepo?.inputBox?.value.trim() ?? '';
}

async function openGitHistory(
    context: vscode.ExtensionContext,
    panelContext: GitHistoryPanelContext = {},
): Promise<void> {
    if (!commitService || !repoDiscovery || !gitActions || !gitRepoCommands || !gitActionHandler) {
        return;
    }
    await repoDiscovery.discover();
    const preferredRepo = resolvePreferredRepo(panelContext, repoDiscovery);
    if (panelContext.fileUri && !preferredRepo) {
        vscode.window.showErrorMessage(i18n('ext.git.notInRepo'));
        return;
    }
    if (preferredRepo) {
        panelContext = { ...panelContext, preferredRepo };
    }
    await GitHistoryPanel.createOrShow(
        context,
        commitService,
        repoDiscovery,
        gitActions,
        gitActionHandler,
        panelContext,
    );
}

async function runQuickSyncCommand(): Promise<void> {
    if (!repoDiscovery || !gitActions) {
        return;
    }
    await repoDiscovery.discover();
    const repos = repoDiscovery.getRepos();
    if (repos.length === 0) {
        vscode.window.showWarningMessage('No Git repository found.');
        return;
    }

    const repo = repos.length === 1
        ? repos[0]
        : await vscode.window.showQuickPick(
            [...repos],
            { title: 'Quick Sync', placeHolder: 'Select repository' },
        );
    if (!repo) return;

    const branch = await gitActions.getCurrentBranch(repo);
    if (!branch) {
        vscode.window.showErrorMessage('Unable to get the current branch.');
        return;
    }
    if (branch === 'HEAD') {
        vscode.window.showWarningMessage('Quick Sync is unavailable in detached HEAD state.');
        return;
    }

    const remotes = await gitActions.listRemotes(repo);
    const remote = remotes.length === 0
        ? ''
        : remotes.length === 1
            ? remotes[0]
            : await vscode.window.showQuickPick(
                remotes,
                { title: 'Quick Sync', placeHolder: 'Select remote' },
            );
    if (remote === undefined) return;

    const hasUncommittedChanges = await gitActions.hasUncommittedChanges(repo);
    let commitMessage = await getGitInputCommitMessage(repo) || 'Quick Sync';
    if (hasUncommittedChanges) {
        const input = await vscode.window.showInputBox({
            title: 'Quick Sync',
            prompt: 'Commit message',
            value: commitMessage,
        });
        if (input === undefined) return;
        commitMessage = input;
    }

    await vscode.window.withProgress(
        { location: vscode.ProgressLocation.Notification, title: 'Git: Quick Sync', cancellable: false },
        async () => {
            const error = await gitActions.quickSync(repo, branch, remote, commitMessage);
            if (error) {
                vscode.window.showErrorMessage(error);
            }
        },
    );
}

export async function activateGitHistory(context: vscode.ExtensionContext): Promise<void> {
    try {
        const gitExecutable = await findGit();
        const executor = new GitExecutor(gitExecutable);
        commitService = new CommitService(executor);
        repoDiscovery = new RepoDiscovery(executor);
        gitActions = new GitActions(executor);
        gitRepoCommands = new GitRepoCommands(executor);
        gitActionHandler = new GitActionHandler(gitRepoCommands);
        repoDiscovery.bindToContext(context);
        await repoDiscovery.discover();
    } catch {
        context.subscriptions.push(
            vscode.commands.registerCommand('excelAiVbaStudio.gitHistory.view', () => {
                vscode.window.showErrorMessage(i18n('ext.git.unableToFindGit'));
            }),
            vscode.commands.registerCommand('excelAiVbaStudio.gitHistory.viewFileHistory', () => {
                vscode.window.showErrorMessage(i18n('ext.git.unableToFindGit'));
            })
        );
        return;
    }

    const statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 50);
    statusBarItem.command = 'excelAiVbaStudio.gitHistory.view';
    statusBarItem.text = '$(git-commit) Git';
    statusBarItem.tooltip = i18n('ext.git.statusBarTooltip');
    statusBarItem.show();
    context.subscriptions.push(statusBarItem);

    context.subscriptions.push(
        vscode.commands.registerCommand('excelAiVbaStudio.gitHistory.view', async (arg?: unknown) => {
            const panelContext = buildPanelContextFromCommandArg(arg);
            await openGitHistory(context, panelContext);
        }),
        vscode.commands.registerCommand('excelAiVbaStudio.gitHistory.viewFileHistory', async (arg?: unknown) => {
            const fromArg = buildPanelContextFromCommandArg(arg);
            const fileUri = resolveFileUri(arg) ?? fromArg.fileUri;
            if (!fileUri) {
                vscode.window.showWarningMessage(i18n('ext.git.openFileFirst'));
                return;
            }
            await openGitHistory(context, mergePanelContext({ fileUri }, fromArg));
        }),
        vscode.commands.registerCommand('excelAiVbaStudio.gitHistory.quickSync', async () => {
            await runQuickSyncCommand();
        }),
        vscode.window.registerWebviewPanelSerializer(
            GIT_HISTORY_VIEW_TYPE,
            new GitHistoryPanelSerializer(context, commitService, repoDiscovery, gitActions, gitActionHandler)
        )
    );
}
