# Third-party notices

Excel AI & VBA Studio incorporates or bundles the open-source components below. Each component remains subject to its own license; nothing in this file changes those terms. Versions refer to release `0.1.0`; `package-lock.json` is the authoritative build inventory.

The complete license texts shipped with the extension are collected in [THIRD_PARTY_LICENSES.txt](THIRD_PARTY_LICENSES.txt).

| Component | Version | License | Source |
| --- | ---: | --- | --- |
| Office Viewer by Weijan Chen | 4.1.7 fork base | MIT | <https://github.com/cweijan/vscode-office> |
| x-data-spreadsheet by myliang | vendored fork | MIT | <https://github.com/myliang/x-spreadsheet> |
| pyOpenVBA | 3.1.0 | MIT | <https://github.com/WilliamSmithEdward/pyOpenVBA> |
| PyInstaller bootloader | 6.15.0 | GPL-2.0-or-later with Bootloader Exception | <https://github.com/pyinstaller/pyinstaller> |
| Python runtime | 3.11.9 | Python Software Foundation License | <https://www.python.org/> |
| Vditor by B3log | source tree only | MIT | <https://github.com/Vanessa219/vditor> |
| SheetJS Community Edition (`xlsx`) | 0.20.3 | Apache-2.0 | <https://git.sheetjs.com/SheetJS/sheetjs> |
| ExcelJS (`@cweijan/exceljs`) | 5.0.1 | MIT | <https://github.com/exceljs/exceljs> |
| React / React DOM | 19.2.8 | MIT | <https://github.com/facebook/react> |
| Ant Design | 5.29.3 | MIT | <https://github.com/ant-design/ant-design> |
| Ant Design Icons | 5.6.1 | MIT | <https://github.com/ant-design/ant-design-icons> |
| VS Code Codicons | 0.0.45 | CC-BY-4.0 | <https://github.com/microsoft/vscode-codicons> |
| JSZip | 3.10.1 | MIT OR GPL-3.0-or-later | <https://github.com/Stuk/jszip> |
| iconv-lite | 0.6.3 | MIT | <https://github.com/ashtuchkin/iconv-lite> |
| uDSV | 0.5.3 | MIT | <https://github.com/leeoniya/uDSV> |
| buffer | 6.0.3 | MIT | <https://github.com/feross/buffer> |

VS Code Codicons and its icon font are Copyright (c) Microsoft Corporation and are used under the [Creative Commons Attribution 4.0 International license](https://creativecommons.org/licenses/by/4.0/). Codicons are bundled without intentional visual modification.

SheetJS Community Edition is used under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0). The pinned distribution contains no separate `NOTICE` file.

Build-only tooling such as Vite, esbuild, TypeScript, and `@vscode/vsce` is not shipped as a runtime dependency in the VSIX.

The bundled Windows x64 VBA write-back helper contains pyOpenVBA 3.1.0. It
edits VBA source in the documented binary container without starting Office.
The helper also embeds the PyInstaller 6.15.0 bootloader under its explicit
Bootloader Exception and the Python 3.11.9 runtime.

The preserved license texts for the Office Viewer baseline, vendored
x-data-spreadsheet code, and the Vditor source tree are stored in the
[LICENSES](LICENSES) directory. Vditor is retained in the public source tree
but is excluded from the targeted Excel AI & VBA Studio VSIX.
