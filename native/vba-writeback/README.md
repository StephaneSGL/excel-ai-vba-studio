# VBA write-back runtime

`cli.py` is the source for the bundled Windows x64 helper. It edits only VBA
source inside `.xlsm`/`.xlam` files. It does not start Excel, execute macros, or
change Office/AccessVBOM registry settings.

Build from a clean Python 3.11 virtual environment:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build-vba-writeback.ps1
```

The build pins the complete Python build toolchain, fixes PyInstaller's
reproducibility inputs, and packages a self-contained console executable.
Windows CI requires a clean rebuild to match the tracked executable
byte-for-byte. The upstream MIT license is included in
`LICENSES/PYOPENVBA-MIT.txt`.

The helper updates existing standard, class, document, and UserForm
code-behind modules. It can add standard or class modules. It deliberately
refuses new document modules, new UserForm designers, signed projects,
password-protected projects, network paths, and reparse points.

The extension serializes writes per workbook, checks the exported workbook
hash, and writes and validates a sibling temporary copy. The atomic replacement
captures the exact displaced workbook as the verified backup; a conflicting
version is restored without blindly overwriting a newer writer. Existing
UserForm designer streams and all non-VBA OOXML payloads are validated after
the patch.

See [`docs/VBA-WRITEBACK-ADR.md`](../../docs/VBA-WRITEBACK-ADR.md) for the
architecture decision, threat model, recovery behavior, and supported limits.
