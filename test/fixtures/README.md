# VBA write-back fixture

`DemoExcelUserForm.xlsm` is the local project workbook supplied for this
repository’s VBA round-trip tests. It contains:

- `mCode` (standard module);
- `ThisWorkbook` and `Data` document modules;
- `oUserForm` with real designer storage and controls.

Automated tests always copy it to a temporary directory before editing. Macros
are never executed.
