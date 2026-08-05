# Graphy codebase analysis — excel-ai-vba-studio

This report is generated from the files tracked by Git. It provides an accessible text equivalent of the image shown in the repository README.

## File structure summary

- Tracked files: **1897**
- Folders: **106**
- File types: **25**
- Source commit: `3e1bc420c9c6`

## File types

| Type | Files |
|---|---:|
| `.svg` | 1310 |
| `.ts` | 245 |
| `.js` | 118 |
| `.tsx` | 47 |
| `.mjs` | 46 |
| `.json` | 31 |
| `.md` | 26 |
| `.css` | 14 |
| `.ps1` | 13 |
| `.less` | 10 |
| `.yml` | 7 |
| `.txt` | 7 |
| `.png` | 6 |
| `no extension` | 2 |
| `.html` | 2 |
| `.py` | 2 |
| `.tmlanguage` | 2 |
| `.xlsm` | 2 |
| `.gitignore` | 1 |
| `.vscodeignore` | 1 |
| `.exe` | 1 |
| `.sh` | 1 |
| `.applescript` | 1 |
| `.lock` | 1 |
| `.xlsx` | 1 |

## Directory tree

```text
excel-ai-vba-studio/
├── .github/  (8 files)
│   ├── ISSUE_TEMPLATE/  (3 files)
│   │   ├── bug_report.yml
│   │   ├── config.yml
│   │   └── feature_request.yml
│   ├── workflows/  (3 files)
│   │   ├── codeql.yml
│   │   ├── main.yml
│   │   └── publish.yml
│   ├── CODEOWNERS
│   └── dependabot.yml
├── .vscode/  (3 files)
│   ├── launch.json
│   ├── settings.json
│   └── tasks.json
├── bin/  (1 files)
│   └── win32-x64/  (1 files)
│       └── excel-ai-vba-writeback.exe
├── docs/  (4 files)
│   ├── CHARTS-AND-TABLES.md
│   ├── PROJECTS.md
│   ├── PUBLISHING.md
│   └── VBA-WRITEBACK-ADR.md
├── icons/  (1263 files)
│   ├── 3d.svg
│   ├── abap.svg
│   ├── abc.svg
│   ├── actionscript.svg
│   ├── ada.svg
│   ├── adobe-illustrator.svg
│   ├── adobe-illustrator_light.svg
│   ├── adobe-photoshop.svg
│   ├── adobe-photoshop_light.svg
│   ├── adobe-swc.svg
│   ├── adonis.svg
│   ├── advpl-include.clone.svg
│   ├── advpl-ptm.clone.svg
│   ├── advpl-tlpp.clone.svg
│   ├── advpl.svg
│   ├── ahk2.clone.svg
│   ├── amplify.svg
│   ├── android.svg
│   ├── angular-component.clone.svg
│   ├── angular-component.svg
│   ├── angular-directive.clone.svg
│   ├── angular-directive.svg
│   ├── angular-guard.clone.svg
│   ├── angular-guard.svg
│   ├── angular-interceptor.clone.svg
│   ├── angular-pipe.clone.svg
│   ├── angular-pipe.svg
│   ├── angular-resolver.clone.svg
│   ├── angular-resolver.svg
│   ├── angular-service.clone.svg
│   ├── angular-service.svg
│   ├── angular.svg
│   ├── antlr.svg
│   ├── apiblueprint.svg
│   ├── apollo.svg
│   ├── applescript.svg
│   └── … 1227 more entries
├── image/  (8 files)
│   ├── README/  (2 files)
│   │   ├── 1666635632251.png
│   │   └── 1783342874748.png
│   ├── README-CN/  (2 files)
│   │   ├── 1685418034035.png
│   │   └── 1711182793554.png
│   ├── git-history-dark.svg
│   ├── git-history-light.svg
│   ├── logo.png
│   └── marketplace-icon.png
├── lib/  (3 files)
│   ├── linux.sh
│   ├── mac.applescript
│   └── pc.ps1
├── LICENSES/  (7 files)
│   ├── OFFICE-VIEWER-MIT.txt
│   ├── POLYFORM-NONCOMMERCIAL-1.0.0.md
│   ├── PYINSTALLER-GPL2-EXCEPTION.txt
│   ├── PYOPENVBA-MIT.txt
│   ├── PYTHON-3.11-PSF.txt
│   ├── VDITOR-MIT.txt
│   └── X-DATA-SPREADSHEET-MIT.txt
├── native/  (3 files)
│   └── vba-writeback/  (3 files)
│       ├── cli.py
│       ├── README.md
│       └── requirements.lock
├── public/  (1 files)
│   └── woff2_decompress_binding.js
├── scripts/  (7 files)
│   ├── apply-vba-designer.ps1
│   ├── inspect-office-security.ps1
│   ├── office-ai-apply-edits.ps1
│   ├── office-ai-export.ps1
│   ├── ooxml-package-signature.ps1
│   ├── open-excel-developer.ps1
│   └── prepare-macro-workbook.ps1
├── snippets/  (1 files)
│   └── http.json
├── src/  (471 files)
│   ├── common/  (17 files)
│   │   ├── vscode-nls-i18n/  (1 files)
│   │   │   └── index.ts
│   │   ├── excelWorkbookObjects.ts
│   │   ├── extensionHost.ts
│   │   ├── extensionResource.ts
│   │   ├── fileReadOnly.ts
│   │   ├── fileSuffix.ts
│   │   ├── fileUtil.ts
│   │   ├── global.ts
│   │   ├── handler.ts
│   │   ├── nativeExcelEdits.ts
│   │   ├── ooxmlPackageSignature.ts
│   │   ├── Output.ts
│   │   ├── reactApp.ts
│   │   ├── simpleEventEmitter.ts
│   │   ├── util.ts
│   │   ├── webviewUri.ts
│   │   └── workspaceFs.ts
│   ├── excelAiVbaStudio/  (13 files)
│   │   ├── explorer.ts
│   │   ├── index.ts
│   │   ├── languageModelTool.ts
│   │   ├── officeSecurity.ts
│   │   ├── security.ts
│   │   ├── securityCenterPanel.ts
│   │   ├── types.ts
│   │   ├── userFormPreview.ts
│   │   ├── vbaInteractionGraph.ts
│   │   ├── vbaStudioPanel.ts
│   │   ├── vbaWritebackService.ts
│   │   ├── workbookObjectTool.ts
│   │   └── workbookService.ts
│   ├── gitHistory/  (23 files)
│   │   ├── provider/  (7 files)
│   │   │   ├── gitActionHandler.ts
│   │   │   ├── gitHistoryPanel.ts
│   │   │   ├── gitHistoryPanelContext.ts
│   │   │   ├── gitHistoryPanelSerializer.ts
│   │   │   ├── index.ts
│   │   │   ├── messageRouter.ts
│   │   │   └── panelHandler.ts
│   │   ├── service/  (7 files)
│   │   │   ├── commitService.ts
│   │   │   ├── findGit.ts
│   │   │   ├── gitActions.ts
│   │   │   ├── gitExecutor.ts
│   │   │   ├── gitRepoCommands.ts
│   │   │   ├── repoDiscovery.ts
│   │   │   └── repoFileWatcher.ts
│   │   ├── types/  (4 files)
│   │   │   ├── git.ts
│   │   │   ├── gitActions.ts
│   │   │   ├── messages.ts
│   │   │   └── repoConfig.ts
│   │   └── util/  (5 files)
│   │       ├── gitHistoryInitPayload.ts
│   │       ├── gitHistoryPreferences.ts
│   │       ├── remoteRefNames.ts
│   │       ├── repoPath.ts
│   │       └── resolveGitHistoryCommandContext.ts
│   ├── provider/  (68 files)
│   │   ├── compress/  (6 files)
│   │   │   ├── commonHandler.ts
│   │   │   ├── decompressHandler.ts
│   │   │   ├── rarHandler.ts
│   │   │   ├── sevenZipHandler.ts
│   │   │   ├── tarHandler.ts
│   │   │   └── zipHandler.ts
│   │   ├── handlers/  (3 files)
│   │   │   ├── imageHanlder.ts
│   │   │   ├── officeContent.ts
│   │   │   └── svgHandler.ts
│   │   ├── http/  (49 files)
│   │   │   ├── common/  (1 files)
│   │   │   │   └── constants.ts
│   │   │   ├── controllers/  (2 files)
│   │   │   │   ├── codeSnippetController.ts
│   │   │   │   └── requestController.ts
│   │   │   ├── models/  (9 files)
│   │   │   │   ├── configurationSettings.ts
│   │   │   │   ├── documentCache.ts
│   │   │   │   ├── harHttpRequest.ts
│   │   │   │   ├── httpElement.ts
│   │   │   │   ├── httpRequest.ts
│   │   │   │   ├── httpResponse.ts
│   │   │   │   ├── httpVariableResolveResult.ts
│   │   │   │   ├── requestParser.ts
│   │   │   │   └── types.ts
│   │   │   ├── providers/  (12 files)
│   │   │   │   ├── customVariableDiagnosticsProvider.ts
│   │   │   │   ├── documentLinkProvider.ts
│   │   │   │   ├── environmentOrFileVariableHoverProvider.ts
│   │   │   │   ├── fileVariableDefinitionProvider.ts
│   │   │   │   ├── fileVariableReferenceProvider.ts
│   │   │   │   ├── fileVariableReferencesCodeLensProvider.ts
│   │   │   │   ├── httpCodeLensProvider.ts
│   │   │   │   ├── httpCompletionItemProvider.ts
│   │   │   │   ├── httpDocumentSymbolProvider.ts
│   │   │   │   ├── requestVariableCompletionItemProvider.ts
│   │   │   │   ├── requestVariableDefinitionProvider.ts
│   │   │   │   └── requestVariableHoverProvider.ts
│   │   │   ├── utils/  (22 files)
│   │   │   │   ├── httpVariableProviders/  (5 files)
│   │   │   │   │   └── … 5 entries below this level
│   │   │   │   ├── combinedStream.ts
│   │   │   │   ├── curlRequestParser.ts
│   │   │   │   ├── httpClient.ts
│   │   │   │   ├── httpElementFactory.ts
│   │   │   │   ├── httpRequestParser.ts
│   │   │   │   ├── mimeUtility.ts
│   │   │   │   ├── misc.ts
│   │   │   │   ├── pretty-data.ts
│   │   │   │   ├── requestParserUtil.ts
│   │   │   │   ├── requestVariableCache.ts
│   │   │   │   ├── requestVariableCacheValueProcessor.ts
│   │   │   │   ├── responseFormatUtility.ts
│   │   │   │   ├── selector.ts
│   │   │   │   ├── streamUtility.ts
│   │   │   │   ├── variableProcessor.ts
│   │   │   │   ├── variableUtility.ts
│   │   │   │   └── workspaceUtility.ts
│   │   │   ├── views/  (1 files)
│   │   │   │   └── responseView.ts
│   │   │   ├── index.ts
│   │   │   └── logger.ts
│   │   ├── xml/  (3 files)
│   │   │   ├── providers/  (1 files)
│   │   │   │   └── xmlDocumentFormattingEditProvider.ts
│   │   │   ├── utils/  (1 files)
│   │   │   │   └── xmlFormat.ts
│   │   │   └── index.ts
│   │   ├── yaml/  (4 files)
│   │   │   ├── providers/  (2 files)
│   │   │   │   ├── yamlDefinitionProvider.ts
│   │   │   │   └── yamlDocumentSymbolProvider.ts
│   │   │   ├── utils/  (1 files)
│   │   │   │   └── yamlDocument.ts
│   │   │   └── index.ts
│   │   ├── archiveViewerProvider.ts
│   │   ├── nativeExcelBridge.ts
│   │   └── officeViewerProvider.ts
│   ├── react/  (295 files)
│   │   ├── i18n/  (11 files)
│   │   │   ├── messages/  (10 files)
│   │   │   │   ├── de.ts
│   │   │   │   ├── en.ts
│   │   │   │   ├── es.ts
│   │   │   │   ├── fr.ts
│   │   │   │   ├── ja.ts
│   │   │   │   ├── ko.ts
│   │   │   │   ├── pt-br.ts
│   │   │   │   ├── ru.ts
│   │   │   │   ├── zh-cn.ts
│   │   │   │   └── zh-tw.ts
│   │   │   └── i18nConfig.ts
│   │   ├── polyfills/  (1 files)
… report truncated; browse the repository for the complete tree
```

The tree is intentionally bounded for readability. GitHub’s **Code** view remains the authoritative complete tree.
