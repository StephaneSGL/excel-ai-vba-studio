# Changelog

All notable changes to Excel AI & VBA Studio are documented here.

The project uses semantic versions. Marketplace Preview status is represented by the `preview` manifest flag; Marketplace versions remain in `major.minor.patch` format.

## [Unreleased]

### Added

- Added a read-only Enterprise Security Center for `.xlsx`, `.xlsm`, `.xls`, and `.xlsb`. It reports file origin/MOTW, VBA and package signatures, EFS and OOXML package encryption, sensitivity-label metadata, macro/VBA settings, ActiveX, Protected View, application-wide and Excel Trusted Locations, the policy controlling user-defined locations, and the detected source of Office 16 policy or preference values.
- Added a per-capability view for the integrated grid, VBA inspection/write-back, UserForm designer, ActiveX creation, and macro execution, with restricted, managed, standard, and unknown workbook levels.
- Added an effective-policy table that keeps Cloud Policy, Windows policy-registry values, and local preferences separate, shows shadowed evidence, and uses the detected Office architecture for registry-view decisions.
- Added separate, non-attributing signals for Microsoft Intune/MDM, Windows Group Policy history, Microsoft 365 Cloud Policy, and Microsoft Purview plus explicit links to the authorized Microsoft 365 Apps, Intune, and Purview admin portals.
- Added structural parsing for the modern OPC `LabelInfo.xml` relationship and the correctly related legacy OOXML custom properties, including technical name, declared tenant ID, assignment method, date, and content-marking bits when locally available.
- Added per-tenant Purview fallback between `LabelInfo` and historical properties, IRMDS container recognition, separate Intune and correlated MDM signals, denied-registry-read states, GPO diagnostic guidance, and legacy XLM policy warnings.

### Changed

- Clarified the noncommercial community-use rights and the separate commercial-license boundary without changing the PolyForm Noncommercial 1.0.0 terms.
- Removed obsolete upstream translations that described a different extension, unsupported features, and telemetry that this project does not collect.
- Excluded local package-output directories consistently from Git.
- Replaced the unsupported Marketplace `--oidc` invocation with a verified VSIX artifact; automatic publication remains intentionally disabled until a Microsoft Entra workload identity is configured for the publisher.

### Security

- The enterprise probe is bounded, local, and read-only: it never opens Excel, runs VBA, removes Mark of the Web, writes the registry, changes Trust Center, or attempts to bypass managed policy.
- The audit now holds a read-only file lock across the complete inspection, parses only one valid `ZoneTransfer` section with `ZoneId` 0–4, expands bounded Trusted Location environment variables, and includes Excel's documented built-in Trusted Locations.
- Macro-capable containers whose VBA or XLM inventory cannot be proven now remain `unknown`; they are never presented as safely macro-free merely because a conventional part path was absent.
- Compound-file traversal is capped by directory-entry count, hierarchy depth, and reconstructed path length; ambiguous structures fail closed.
- GitHub Actions are pinned to immutable commits, CodeQL scans JavaScript/TypeScript and Python, native Python build wheels are pinned by SHA-256, and Windows CI rebuilds, hashes, and tests the native helper before a tagged package consumes that exact artifact. Package validation uses an explicit allowlist.
- Removed the invalid cross-host byte-for-byte PyInstaller comparison; runner-built executables are now identified by digest and accepted only after the native security and atomic-write suites pass.
- Cloud, machine, user, and preference sources remain visible independently; unknown effective decisions are reported as unknown instead of being treated as permission.
- A Windows policy-registry value is never attributed to GPO or Intune from enrollment presence alone, and local Purview metadata is never treated as authenticated tenant policy or proof of encryption.
- The Center never opens the inspected workbook, rejects malformed policy value types, bounds Mark-of-the-Web and enrollment reads, and reports unreadable higher-priority policy stores as unknown instead of falling back to a weaker preference.
- Every workbook mutation path now resolves OOXML package signatures through OPC origin/signature relationships and effective Content Types, including arbitrary valid part URIs; malformed, orphaned, external, unreadable, or otherwise ambiguous signature state fails closed.

### Tests

- Added static guards against registry/file/network writes and Excel automation plus synthetic workbook probes for bounded JSON, modern and legacy Purview metadata, `EnabledV2` precedence, orphan rejection, Intune non-attribution, policy precedence, registry bitness, source typing, and Trusted Location limits.
- Added regressions for invalid and ambiguous Mark-of-the-Web streams, macro-capable containers with inconclusive inventories, built-in Excel Trusted Locations, deep compound-file hierarchies, and Linux-safe package-signature validation.
- Added an executable signed-package regression covering XLSX/XLSM grid writes, Save As, VBA bootstrap/write-back, UserForms, worksheet buttons, and ActiveX without running a macro.

## [0.5.1] - 2026-07-28

### Fixed

- Fixed `createWorksheetButton` and `assignWorksheetButtonMacro` verification when Excel normalizes workbook or VBA identifier casing in `OnAction`. Verification still requires the same workbook and macro target.
- Kept machine-readable PowerShell export JSON in the output log instead of replacing the user-facing export progress message with the raw payload.
- Removed the XLSX-to-XLSM bootstrap dependency on PowerShell module autoloading by using a self-contained SHA-256 implementation.

### Tests

- Added live Excel coverage using a deliberately case-varied qualified macro name and static guards against case-sensitive `OnAction` verification.

## [0.5.0] - 2026-07-28

### Added

- Added a visual UserForm designer inside VBA Studio, backed by the guarded Excel COM inventory of real forms and controls. Standard controls can be added, selected, moved on a four-point grid, resized, and edited through a property inspector.
- Added `updateUserFormControl` to `#excelVbaDesignWorkbook` so GitHub Copilot can reposition or update existing UserForm controls transactionally.
- Added `setUserFormEventHandler` for complete `Private Sub object_event(...)` procedures, including complex event signatures with MSForms parameters.
- Added a visual event editor for UserForm and control events. Existing handlers require the explicit `replaceExisting=true` opt-in and a modal confirmation from the visual editor.

### Security

- Event writes accept one bounded, exact event procedure only. They preserve every unrelated procedure, never execute VBA, verify the persisted source in a second Excel instance, and retain the existing atomic backup and rollback guarantees.
- UserForm geometry comes from the real VBA designer object model rather than inferred workbook serialization.

### Tests

- Added live Excel coverage for control geometry updates, parameterized `KeyDown`, `UserForm_Initialize`, explicit event replacement, replacement refusal, persisted-source verification, and process cleanup.
- Added a native UserForm inventory regression that verifies all 12 fixture controls and their real geometry.

## [0.4.0] - 2026-07-28

### Added

- Added a VBA Controls view that inventories worksheet Form Control buttons and worksheet ActiveX controls during the existing guarded Excel export pass.
- Added a static, non-executing button-to-macro-to-UserForm graph, macro reassignment for existing Form Control buttons, and verified `Click` binding for supported worksheet ActiveX buttons.
- Added bounded worksheet ActiveX creation for built-in MSForms controls. Third-party controls require an exact ProgID already present in the user-owned `excelAiVbaStudio.allowedCustomActiveXProgIds` allowlist.
- Added an isolated VBA Developer workspace opened in a separate VS Code window, with an extension-owned marker, path and hash validation, exported sources, workbook metadata, and automatic VBA Studio opening.

### Fixed

- Fixed false `worksheet-features` rejections when the embedded grid normalized existing images, page setup, links, or other read-only sheet metadata during `.xlsm` loading. The native edit baseline is now captured after grid hydration, while actual unsupported edits remain refused.
- Routed real UserForm, worksheet-button, and worksheet-ActiveX operations through the transactional VBA Designer instead of generic workbook serialization.

### Security

- Macro relationships are parsed statically; the VS Code simulation never executes VBA.
- ActiveX insertion remains subject to the user's Excel Trust Center and organization policy. The extension never changes ActiveX, AccessVBOM, registry, or macro-security settings.
- Existing ActiveX event handlers are never overwritten, and custom ProgIDs are denied by default.

### Tests

- Added graph-resolution tests for Form Controls and ActiveX `Click` handlers, plus a regression test for feature-rich `.xlsm` sheets that receive supported cell edits.
- Extended live Excel tests with Form Control reassignment, ActiveX creation and binding where Office permits it, custom-ProgID allowlisting, policy-blocked insertion refusal, rollback, native inspection, and process-leak checks.

## [0.3.0] - 2026-07-28

### Added

- Added `#excelVbaDesignWorkbook` for creating complete UserForms with real designer/`.frx` streams, adding supported visual controls, and creating worksheet Form Control buttons with macro assignments in existing `.xlsm` workbooks.
- Added an interactive UserForm preview for standard controls, with explicit native Excel/VBE handoff and no VBA execution inside VS Code.
- Added native `.xlsm` persistence for explicit row heights, column widths, the five conditional-formatting presets exposed by the ribbon, and explicit worksheet-wide conditional-rule clearing.

### Fixed

- Kept fake `.frm` creation refused in the source-only writer and routed real UserForm creation through the separate verified VBA Designer transaction.
- Refused worksheet-button creation unless its `OnAction` target is an existing public, zero-argument standard-module macro.
- Preserved worksheet names containing dots during native button verification.
- Made normal **Open in Excel** return immediately after the verified launch request; exact-window polling remains limited to native VBE handoff.

### Security

- The VBA Designer rejects stale, signed, protected, network, reparse-point, and oversized requests; uses same-directory staging; verifies native UserForm streams; preserves the displaced original; and never runs a macro or changes AccessVBOM.
- A post-replacement failure restores the verified displaced workbook while retaining its persistent recovery backup.
- Native formatting operations use a strict allowlist, reject arbitrary formulas and ambiguous resets, append rules without rebuilding existing ones, and retain the existing VBA/OOXML fingerprint and rollback transaction.

### Tests

- Added live Excel coverage for all 12 supported standard UserForm controls, new designer/`.frx` streams, controls on existing forms, verified worksheet-button macro assignment, exact backup hashes, native post-inspection, stale-write rollback, and absence of leaked Excel processes.
- Extended live Excel coverage to row/column dimensions, all five supported conditional-rule families, full rule clearing, exact backups, button/opaque-part preservation, and absence of leaked Excel processes.

## [0.2.1] - 2026-07-27

### Added

- Added first-write `.xlsx` to sibling `.xlsm` conversion for standard modules and classes through `#excelVbaWriteModule`.
- Returned the exact `targetWorkbookPath` so Copilot continues every subsequent write against the created macro-enabled workbook.

### Fixed

- Prevented Copilot from presenting exported `.bas` or `.frm` working-copy files as successful workbook modifications.
- Replaced the misleading dynamic-UserForm fallback with an explicit refusal: a real new UserForm still requires the native VBE designer and `.frx`.
- Corrected stale English and French warnings so only legacy `.xls`, not editable `.xlsm`, is described as read-only.

### Security

- Refused every write path containing the exact `.excel-ai-vba-backups` component.
- The conversion keeps the `.xlsx` source unchanged, disables macros/events/links, owns and releases the exact hidden Excel process, never changes AccessVBOM, and verifies the persisted `xl/vbaProject.bin`.

### Tests

- Added live Excel coverage for `.bas` and `.cls` persistence, unchanged `.xlsx` hashes, backup-path rejection, and absence of leaked Excel processes.

## [0.2.0] - 2026-07-26

### Added

- Added transactional VBA source write-back for local `.xlsm` and `.xlam` workbooks.
- Added the prompt-referenceable `#excelVbaWriteModule` tool for GitHub Copilot.
- Added automatic reinjection when an exported `.bas`, `.cls`, or existing `.frm` source is saved in VS Code.
- Added safe creation of standard and class modules from the VBA Studio or Copilot.
- Bundled a self-contained Windows x64 pyOpenVBA helper and a real UserForm regression fixture.
- Added targeted native `.xlsm` editing for cell values, formulas, number
  formats, fonts, alignment, fills, and borders without leaving VS Code.
- Restored explicit commands and ribbon actions for opening the real workbook
  in Microsoft Excel and opening Excel's native VBE.
- Added the VS Code editable-document lifecycle, including close prompts,
  revert, Save As routing, and bounded hot-exit recovery for unsaved grids.

### Fixed

- Removed the embedded read-only VBA banner for writable `.xlsm` workbooks;
  legacy `.xls` files remain protected.
- Made the XLSX reader accept namespaced SpreadsheetML elements and absolute
  package relationship targets.
- Prevented file-watcher refreshes from overwriting unsaved grid edits.
- Prevented split-view, edit-during-save, and stale hot-exit recovery races
  from clearing or overwriting newer spreadsheet changes.
- Routed Save, Save As, and external revert through the matching active custom
  editor instead of issuing a delayed global save command.

### Security

- Every changed workbook receives the exact atomically displaced source as a
  verified timestamped backup.
- Native cell edits now use a hashed sibling work copy, re-check the source
  immediately before commit, validate the macro-enabled OOXML package, retain a
  persistent backup, and atomically replace the original.
- Native saves bind the grid snapshot to a writer-locked source, fingerprint
  every logical VBA/UserForm stream, and restore the verified backup
  automatically if post-replacement validation fails.
- The bundled writer uses a fully pinned Python package toolchain. Windows CI
  rebuilds it from source, records its digest, and runs the native regression
  suite against the resulting executable.
- Stale workbooks, signed or password-protected VBA projects, reparse points, network paths, oversized requests, and UserForm designer changes are refused.
- The write-back engine never starts Excel, executes macros, edits the Office registry, or changes AccessVBOM.

### Tests

- Added round-trip coverage for standard modules, new modules, existing UserForm code-behind, unchanged designer streams, stale-write rejection, OOXML preservation, and absence of new Excel processes.
- Added real Excel integration coverage for targeted XLSM edits, conditional
  formatting, form buttons, OOXML parts, and bit-for-bit VBA/UserForm streams.

## [0.1.7] - 2026-07-25

### Added

- Added an in-editor conditional-formatting dialog for value rules, text rules, colour scales, data bars, and icon sets.
- Activated the former placeholder ribbon actions through direct spreadsheet operations or a targeted GitHub Copilot request with workbook and VBA context.
- Added in-editor data import, table formatting, deduplication, text-to-columns, subtotals, forecasts, page setup, formula auditing, image insertion, and generated chart/shape visuals.

### Fixed

- Made conditional-format icon sets reliable in VS Code by rendering explicit coloured glyphs instead of platform-dependent colour emoji.
- Removed every visible “Coming soon”/“Bientôt” badge from the ribbon.

## [0.1.6] - 2026-07-25

### Performance

- Reduced the initial webview JavaScript from about 2.00 MB to 219 KB by loading the spreadsheet interface progressively.
- Deferred Excel export, CSV parsing, and legacy XLS/ODS engines until those features are actually used.
- Loaded the workbook parser and sort-state reader concurrently instead of opening the XLSX archive sequentially twice.
- Removed a redundant full worksheet traversal and skipped unnecessary formatting snapshots for XLSX/XLSM files.
- Added an automated bundle-size and lazy-loading regression check.

## [0.1.5] - 2026-07-25

### Changed

- The spreadsheet and VBA Studio now follow the active VS Code light, dark, or high-contrast color theme automatically.
- Removed the saved light-mode override that could conflict with the theme selected by the user in VS Code.
- Improved VBA Studio surfaces, selections, controls, focus rings, warnings, and status colors in light and dark themes.

## [0.1.4] - 2026-07-25

### Added

- Added a complete VBE-style editor tab inside VS Code with project tree, properties, module/procedure selectors, line numbers, source editing, and module/class/UserForm creation.
- Connected the integrated editor to real `.bas`, `.cls`, and `.frm` workspace files so GitHub Copilot and the native VS Code editor share the same sources.
- Added an in-studio Copilot action backed by the generated workspace instructions and `#excelVbaWorkbook`.
- Added a starter `Module1.bas` for `.xlsx` workbooks that do not yet contain an embedded VBA project.

### Fixed

- Kept the Copilot-indexed VBA directory stable during refresh, fixing Windows `EBUSY: resource busy or locked, rmdir ...\\vba`.

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
- Initial Marketplace OIDC workflow with tag/version verification; this path
  was not operational with VSCE and is superseded by the verified-artifact
  workflow documented under Unreleased.

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
- The initial publishing design avoided a stored Marketplace PAT, but its OIDC
  command was unsupported and did not provide a working publication path.
