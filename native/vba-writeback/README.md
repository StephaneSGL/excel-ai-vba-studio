# VBA write-back runtime

`cli.py` is the source for the bundled Windows x64 helper. It edits only VBA
source inside `.xlsm`/`.xlam` files. It does not start Excel, execute macros, or
change Office/AccessVBOM registry settings.

Build from a clean Python 3.11 virtual environment:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build-vba-writeback.ps1
```

The build pins `pyOpenVBA==3.1.0` and packages a self-contained console
executable with PyInstaller. The upstream MIT license is included in
`LICENSES/PYOPENVBA-MIT.txt`.

The helper updates existing standard, class, document, and UserForm
code-behind modules. It can add standard or class modules. It deliberately
refuses new document modules, new UserForm designers, signed projects,
password-protected projects, network paths, and reparse points.

The extension serializes writes per workbook, checks the exported workbook
hash, writes and validates a sibling temporary copy, creates a verified backup,
then performs an atomic replacement. Existing UserForm designer streams and
all non-VBA OOXML payloads are validated after the patch.
