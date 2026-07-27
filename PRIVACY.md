# Privacy statement

Effective date: 2026-07-27

Excel AI & VBA Studio is designed for local workbook processing.

## Data the extension processes

When requested by the user, the extension may read:

- spreadsheet values, formulas, display formats, names, tables, charts,
  comments, validations, links, and workbook metadata;
- connection and Power Query metadata exposed by Microsoft Excel;
- VBA project metadata and module source;
- local file paths needed to open the selected workbook.

The direct VBA writer reads and writes permitted VBA source locally without
enabling AccessVBOM or changing Office/registry settings.

## Local writes and storage

The extension writes a workbook only after an explicit user save.

- `.xlsm` cell changes use the dedicated, targeted Excel writer.
- permitted `.xlsm`/`.xlam` VBA source changes use the direct VBA writer.
- unsafe or unsupported macro-package changes are refused.
- legacy `.xls` files remain read-only.

Each successful targeted write creates a verified backup in
`.excel-ai-vba-backups` beside the workbook. A backup contains the complete
workbook, including its data, formulas, formatting, macros, UserForms, and
other embedded content. It remains until the user removes it.

Generated Markdown, JSON, manifests, and permitted VBA exports are stored in
extension-controlled local storage. They remain until removed by the user,
VS Code, extension uninstallation, or the **Nettoyer les exports générés**
command. Transaction work copies are removed after the operation.

While a spreadsheet has unsaved grid edits, VS Code may create a hot-exit
recovery backup in its extension storage. This local JSON backup contains the
current sheet state and the SHA-256 of the exact source loaded by the editor.
Recovery is refused if the source file has since changed, preventing an older
backup from overwriting newer disk data. The backup is deleted by VS Code after
save, revert, or when it is no longer needed.

The extension does not collect analytics or telemetry. It does not create an
advertising identifier, user profile, extension-managed cloud account, or
extension-managed API key.

## Network and AI providers

The extension does not upload a workbook or generated context to a remote
service by itself.

If the user explicitly invokes `#excelVbaWorkbook`, uses a VBA language-model
tool, pastes an export into chat, or otherwise submits generated context to a
VS Code AI feature, the selected AI/model provider receives that submitted
content. Its privacy policy, retention rules, organization controls, and
service terms then apply.

Microsoft Excel, Visual Studio Code, installed add-ins, selected AI providers,
organization security software, and operating-system services are separate
products with their own data practices.

## User controls

Users can:

- limit exported rows and columns in extension settings;
- disable VBA inclusion;
- review generated Markdown/JSON/VBA files before sharing them;
- remove generated exports and workbook backups;
- save, revert, or close an edited grid to let VS Code retire hot-exit backups;
- use native Excel or the VBE when the safe writer refuses an operation;
- uninstall the extension.

Never share workbooks, backups, or exports containing secrets, personal data,
regulated information, or confidential business information unless the
receiving service and organization policy explicitly permit it.
