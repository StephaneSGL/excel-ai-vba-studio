# Changelog

All notable changes to Excel AI & VBA Studio are documented here.

The project uses semantic versions. Marketplace Preview status is represented by the `preview` manifest flag; Marketplace versions remain in `major.minor.patch` format.

## [0.1.1] - 2026-07-24

### Documentation

- Rebuilt the public README with installation instructions, a format support matrix, architecture, security boundaries, settings, roadmap, and support links.
- Added an explicit licensing map, upstream baseline, CODEOWNERS, and clearer contribution rules.
- Added missing preserved notices for the vendored x-data-spreadsheet code and the Vditor source tree.

### Licensing

- Project-specific contributions distributed from version 0.1.1 are source-available under PolyForm Noncommercial 1.0.0.
- The Office Viewer baseline and all third-party portions retain their original licenses.
- Earlier MIT grants remain valid for versions and commits that were already distributed under MIT.

## [0.1.0] - 2026-07-24

### Added

- Focused Windows x64 extension identity for Excel, CSV, VBA, and AI-assisted workbook inspection.
- Embedded Excel/CSV custom editor.
- Native Microsoft Excel and Developer/VBE launch commands.
- Local Markdown and JSON workbook-context exporter.
- Excel & VBA sidebar explorer.
- Prompt-referenceable `#excelVbaWorkbook` language-model tool.
- Local export cleanup and bounded export settings.
- Marketplace OIDC publishing workflow with tag/version verification.

### Changed

- Marketplace contributions and the bundled runtime are limited to Excel/CSV and Excel/VBA/AI commands.
- SheetJS Community Edition is sourced from the official SheetJS CDN package.
- Extension telemetry, sponsor integrations, and unrelated Office viewers have been removed.
- `.xlsm` and `.xls` now show an explicit VBA-protection banner with direct actions for native Excel and the VBE; embedded Save and Save As are blocked.

### Security

- Context export refuses to open a workbook unless macro automation can be forced off.
- The extension never enables Excel’s programmatic VBA-project access setting.
- VBA and legacy macro formats are read-only in the embedded grid until safe VBA preservation is available.
- Generated context is confined to extension-controlled local storage.
- Publishing uses short-lived OIDC credentials instead of a stored Marketplace PAT.
