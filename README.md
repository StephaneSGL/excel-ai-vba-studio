<p align="center">
  <img src="image/marketplace-icon.png" width="128" alt="Excel AI & VBA Studio">
</p>

<h1 align="center">Excel AI & VBA Studio</h1>

<p align="center">
  <strong>Work with Excel workbooks, VBA projects, and bounded AI context directly in Visual Studio Code.</strong>
</p>

<p align="center">
  <a href="https://github.com/StephaneSGL/excel-ai-vba-studio/actions/workflows/main.yml"><img src="https://github.com/StephaneSGL/excel-ai-vba-studio/actions/workflows/main.yml/badge.svg?branch=main" alt="Validation"></a>
  <img src="https://img.shields.io/badge/status-Preview-f59e0b" alt="Preview">
  <img src="https://img.shields.io/badge/platform-Windows%20x64-0078d4" alt="Windows x64">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-PolyForm%20Noncommercial-7a2f8f" alt="PolyForm Noncommercial"></a>
</p>

Excel AI & VBA Studio is a preview VS Code extension for inspecting supported spreadsheet formats, viewing formulas and values, and working with Excel VBA modules, classes, and UserForms without leaving VS Code.

## Current capabilities

| Format | Workbook grid | VBA workspace | Bounded AI context |
| --- | --- | --- | --- |
| `.xlsx` | Read and targeted edit | Working project without embedded macros | Yes |
| `.csv`, `.tsv` | Read and edit | Not applicable | Yes |
| `.xlsm` | Targeted cell edit | Yes, when the Excel strategy allows it | Yes |
| `.xls` | Protected read-only view | Yes, when the Excel strategy allows it | Yes |
| `.xlsb` | No integrated grid | Yes, when the Excel strategy allows it | Yes |

- Integrated workbook grid inside VS Code.
- Formula, value, style, table, chart, validation, comment, and metadata exports within explicit bounds.
- Excel and VBA project explorer for Microsoft Excel objects, modules, classes, UserForms, and references.
- VBA Studio tab with project explorer, properties, code, procedures, and component creation.
- `.bas`, `.cls`, and `.frm` editing in VS Code.
- Transactional reinjection of supported VBA source into a working `.xlsm` copy.
- Explicit commands for opening the real Excel workbook or native VBE.
- AI tools `#excelVbaWorkbook` and `#excelVbaWriteModule`, invoked only on request.
- No extension telemetry and no API key management by the extension.

## Install from a VSIX

```powershell
code --install-extension .\excel-ai-vba-studio-win32-x64-<version>.vsix
```

You can also use **Extensions → … → Install from VSIX** in VS Code.

## Build from source

Requirements: Node.js 22, npm, Git, and Visual Studio Code 1.95 or newer.

```powershell
git clone https://github.com/StephaneSGL/excel-ai-vba-studio.git
Set-Location excel-ai-vba-studio
npm ci
npm run validate
npm run package
```

The Visual Studio Marketplace release is not available yet. This public repository is currently the official source for the preview.

## Quick start

1. Open a supported workbook in VS Code.
2. Use the integrated grid or the **Excel & VBA** explorer.
3. Export local context only when you need to inspect it.
4. Reference `#excelVbaWorkbook` explicitly from a compatible VS Code AI chat.
5. Open **VBA Studio** to view the real project sources.
6. Edit and save supported VBA source; the extension validates the working copy and keeps a backup before replacement.

## Security and privacy

- Excel is controlled through a dedicated working copy for supported `.xlsm` operations.
- The extension does not change Excel's **Trust access to the VBA project object model** setting.
- Signed or protected VBA projects, network paths, and UserForm designer changes are refused by the reinjection engine.
- The engine never starts a macro during source reinjection.
- Workbook content is treated as untrusted data, not as instructions for an AI model.
- No workbook is sent to an AI provider automatically.

This is not a network sandbox: Microsoft Excel, Windows, installed add-ins, and security software may have their own network behavior. Read [SECURITY.md](SECURITY.md) and [PRIVACY.md](PRIVACY.md) before using real professional workbooks.

## Preview limitations

- Windows x64 only.
- Microsoft Excel desktop is required for COM-based VBA extraction and legacy formats.
- Protected, corrupted, or enterprise-restricted workbooks may provide partial context.
- Integrated `.xlsm` editing is limited to supported values, formulas, and cell styles.
- Creating a complete new UserForm, designer, `.frx`, or sheet button still requires the native VBE.

## License

This project is licensed under the PolyForm Noncommercial License. See [LICENSE](LICENSE) for the exact terms.
