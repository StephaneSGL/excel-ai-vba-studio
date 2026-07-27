# VBA write-back fixture

`DemoExcelUserForm.xlsm` is the local project workbook supplied for this
repository’s VBA round-trip tests. It contains:

- `mCode` (standard module);
- `ThisWorkbook` and `Data` document modules;
- `oUserForm` with real designer storage and controls.

Automated tests always copy it to a temporary directory before editing. Macros
are never executed.

`NativeEditingSynthetic.xlsm` exercises targeted value, formula, number-format,
font, alignment, fill, border, conditional-formatting, and form-button
preservation through hidden Excel automation. SHA-256:
`1681D63080BB902EB8821FC4DDA6F4919630593A88DF846D00DAAFDDC04682A1`.

## XLSX namespace-prefix fixture

`Excel-AI-VBA-Studio-Demo-base.xlsx` is the exact generated workbook that
exposed prefixed SpreadsheetML elements (`x:workbook`, `x:worksheet`, and
related elements) plus absolute package relationship targets. The reader
regression test opens this fixture in memory and verifies its five worksheets.
The fixture is never rewritten.
