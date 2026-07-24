# Excel AI & VBA Studio

> Preview — Windows x64

Open Excel and CSV files in Visual Studio Code, inspect workbook structure with AI, and jump to the real Microsoft Excel or VBA interface when you need the full native toolset.

Visualisez les fichiers Excel et CSV dans Visual Studio Code, fournissez à l’IA un contexte détaillé du classeur, puis ouvrez le véritable Excel ou l’éditeur VBA pour les fonctions natives avancées.

## Features / Fonctionnalités

- Embedded viewer/editor for `.xlsx`, `.csv`, and `.tsv`; `.xlsm` and legacy `.xls` remain view-only in VS Code so the embedded writer can never strip VBA projects or legacy workbook records.
- Native Microsoft Excel launch with the complete ribbon, including Data and Developer tools.
- Direct access to the Visual Basic Editor (VBE) and the workbook VBA project explorer.
- Local, bounded workbook export in Markdown and JSON: values, formulas, formats, tables, charts, names, links, validations, comments, connections, Power Query metadata, and VBA metadata when permitted.
- **Excel & VBA** explorer in the VS Code sidebar.
- VS Code language-model tool `#excelVbaWorkbook` for explicitly sharing workbook context with an AI assistant.
- No extension telemetry and no extension-managed API key.

## Requirements

- Windows x64.
- Visual Studio Code 1.95 or later.
- Microsoft Excel desktop is required for native Excel, VBA, legacy workbook, and COM-based context-export features.
- The embedded viewer does not reproduce every proprietary Excel feature. Commands that need the complete ribbon, Power Query, data tools, add-ins, macros, or VBE open the installed Microsoft Excel application.

## Quick start / Démarrage rapide

1. Open a supported spreadsheet in VS Code.
2. Use the editor toolbar or the **Excel & VBA** explorer.
3. Run one of these commands from the Command Palette:

| Command | Purpose |
| --- | --- |
| `Excel AI & VBA Studio : Ouvrir dans Microsoft Excel` | Open the active file in native Excel. |
| `Excel AI & VBA Studio : Ouvrir Excel en mode Développeur / VBA` | Open Excel and display the Developer/VBE interface. |
| `Excel AI & VBA Studio : Exporter le contexte du classeur` | Create local Markdown and JSON context files. |
| `Excel AI & VBA Studio : Exporter et copier le contexte` | Export and copy a bounded context to the clipboard. |
| `Excel AI & VBA Studio : Exporter et révéler les sources VBA` | Reveal exported VBA files when Excel grants access. |
| `Excel AI & VBA Studio : Nettoyer les exports générés` | Remove extension-generated context exports. |

Keyboard shortcuts:

- `Ctrl+Alt+E`: open in native Excel.
- `Ctrl+Alt+F11`: open Excel in Developer/VBA mode.

In an AI chat that supports VS Code language-model tools, reference `#excelVbaWorkbook`. The workbook is read only when the tool is explicitly invoked or an export command is run.

## VBA security

Displaying the VBE does not require the extension to enable unsafe settings. Reading VBA module source programmatically is different: Excel blocks it unless the user explicitly enables **Trust access to the VBA project object model** in Excel Trust Center.

The extension never changes that security setting. Macro execution is disabled during context export. Only enable VBA project access when you understand and accept the security implications.

The embedded grid deliberately blocks editing, direct save, and Save As for `.xlsm` and `.xls`. Use **Ouvrir dans Microsoft Excel** or **Ouvrir Excel en mode Développeur / VBA** to edit those formats without rebuilding the workbook.

## Data and privacy

Workbook processing and generated exports are local. The extension contains no telemetry client and does not send workbooks to a service by itself. When you deliberately pass generated context to an AI feature, the selected VS Code AI/model provider processes that context under its own terms.

See [PRIVACY.md](PRIVACY.md) for the complete data-handling statement and [SUPPORT.md](SUPPORT.md) before reporting an issue.

## Preview limitations

- This first release targets Windows x64 only.
- Microsoft Excel must be installed for native Excel/VBA integration.
- `.xlsb` can be handled by native Excel/context commands but is not rendered by the embedded editor.
- Password-protected, corrupted, or policy-restricted workbooks may expose only partial metadata.
- VBA source export depends on Excel Trust Center policy; the VBE can still be opened visually when source export is blocked.

## Open-source notice

This project is a focused fork derived from **Office Viewer** by Weijan Chen and remains available under the MIT License. See [NOTICE.md](NOTICE.md), [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md), and [LICENSE](LICENSE).
