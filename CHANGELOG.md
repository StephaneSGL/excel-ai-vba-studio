# Changelog

All notable changes to Excel AI & VBA Studio are documented here.

The project uses semantic versions. Marketplace Preview status is represented by the `preview` manifest flag; Marketplace versions remain in `major.minor.patch` format.

## [0.1.3] - 2026-07-25

### Added

- Added theme, indexed-colour, and tint resolution so workbook colours survive embedded viewing and `.xlsx` saves.
- Added conditional-format rendering and round-trip support for cell rules, colour scales, data bars, and icon sets.
- Added comments, workbook statistics, sheet protection, safe regional sorting, and conditional-format creation to the ribbon.
- Added a direct Copilot action backed by `#excelVbaWorkbook`, plus a dedicated action to export and open VBA sources in VS Code.
- Added a VBE-style VS Code project tree for Excel objects, UserForms, standard modules, class modules, and references.
- Added a synchronized VBA properties view and an internal UserForm preview.
- Added a Copilot-indexable VBA workspace root with generated repository instructions.
- Added native VBComponent export so UserForm `.frm` layout metadata and `.frx` resources are preserved when available.
- Kept unfinished ribbon commands inside VS Code instead of opening Microsoft Excel.

### Fixed

- Fixed blank print previews caused by formatted empty rows and corrected A4 portrait pagination.
- Preserved cached results for structured formulas that the embedded calculation engine cannot evaluate.
- Fixed VBA extraction when VS Code exposes global extension storage through a non-file URI.

## [0.1.1] - 2026-07-24

### Added

- Added a responsive Excel-style ribbon with File, Home, Insert, Page Layout, Formulas, Data, Review, View, and AI & VBA tabs.
- Connected supported ribbon actions to the embedded spreadsheet engine, including formatting, formulas, filters, freeze panes, clipboard, row/column operations, validation, printing, and find/replace.
- Added direct ribbon actions for native Excel, the VBA editor, workbook-context export, and VBA source export.

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
