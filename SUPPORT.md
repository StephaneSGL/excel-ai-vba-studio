# Support

Excel AI & VBA Studio is currently a Preview release.

## Before reporting a problem

Please confirm:

1. The operating system is Linux x64 or Windows x64, and VS Code is version 1.95 or later.
2. The installed VSIX matches the operating system. Linux uses the portable spreadsheet mode; it does not require Microsoft Excel or PowerShell.
3. On Windows, Microsoft Excel Desktop is installed when using native Excel, VBA, UserForms, tables/charts or context-export commands.
4. The latest extension version is installed.
5. The issue also occurs after reopening the synthetic test workbook. For Windows native tools, also try **Excel AI & VBA Studio : Actualiser l’explorateur**.

Portable mode intentionally opens macro-enabled, signed and unsupported complex workbooks read-only. Include the displayed preservation reason when reporting an unexpected refusal. The extension supplies an Excel-like grid; it does not implement every Excel function, VBA runtime or Office integration on Linux.

## Information to include

- Extension, VS Code, operating system and architecture; Microsoft Excel version when using native Windows features.
- Workbook extension (`.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`, or `.tsv`).
- The exact command used and complete error message.
- Minimal reproduction steps.
- Whether the workbook uses macros, protection, external links, Power Query, or organization security policies.

Do not upload confidential workbooks, generated context, VBA source, credentials, personal information, or company data. Prefer a synthetic workbook that reproduces the issue.

## Get help

- Ask usage questions and suggest ideas in [GitHub Discussions](https://github.com/StephaneSGL/excel-ai-vba-studio/discussions).
- Report reproducible defects with a synthetic workbook in [GitHub Issues](https://github.com/StephaneSGL/excel-ai-vba-studio/issues).
- Check existing [issues](https://github.com/StephaneSGL/excel-ai-vba-studio/issues) and [discussions](https://github.com/StephaneSGL/excel-ai-vba-studio/discussions) before opening a duplicate.

## Security reports

Do not disclose a suspected vulnerability in a public issue or discussion. Use the repository’s [private vulnerability reporting](https://github.com/StephaneSGL/excel-ai-vba-studio/security/advisories/new). If that form is unavailable, use the private security contact shown in the Marketplace publisher profile.

Support is provided on a best-effort basis and has no guaranteed response time.
