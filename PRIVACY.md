# Privacy statement

Effective date: 2026-07-24

Excel AI & VBA Studio is designed for local workbook processing.

## Data the extension processes

When requested by the user, the extension may read:

- spreadsheet values, formulas, display formats, names, tables, charts, comments, validations, links, and workbook metadata;
- connection and Power Query metadata exposed by Microsoft Excel;
- VBA project metadata and module source only when Excel policy explicitly permits programmatic access;
- local file paths needed to open the selected workbook.

## Storage

Generated Markdown, JSON, manifest, and permitted VBA exports are stored in extension-controlled local storage. They remain on the computer until removed by the user, VS Code, extension uninstallation, or the **Nettoyer les exports générés** command.

The extension does not collect analytics or telemetry. It does not create an advertising identifier, user profile, or extension-managed cloud account.

The embedded viewer never writes `.xlsm` or `.xls` files. Those formats are view-only and must be edited in native Microsoft Excel to avoid removing macros or legacy workbook records.

## Network and AI providers

The extension does not upload a workbook or generated context to a remote service by itself and does not require an extension-managed API key.

If the user explicitly invokes `#excelVbaWorkbook`, pastes an export into chat, or otherwise submits generated context to a VS Code AI feature, that selected AI/model provider receives the submitted context. Its privacy policy, retention rules, organization controls, and service terms then apply. Users are responsible for checking those terms before sharing sensitive material.

Microsoft Excel, Visual Studio Code, installed add-ins, organization security software, and operating-system services are separate products with their own data practices.

## User controls

Users can:

- limit exported rows and columns in extension settings;
- disable VBA inclusion;
- review Markdown/JSON files before sharing them;
- delete generated exports from the Command Palette;
- uninstall the extension.

Never share workbooks or generated exports containing secrets, personal data, regulated information, or confidential business information unless the receiving service and organization policy explicitly permit it.
