import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputPath = path.join(root, "THIRD_PARTY_LICENSES.txt");
const normalizeLineEndings = (text) => text.replace(/\r\n?/g, "\n");

const components = [
  ["x-data-spreadsheet vendored code", "LICENSES/X-DATA-SPREADSHEET-MIT.txt"],
  ["pyOpenVBA (3.1.0)", "LICENSES/PYOPENVBA-MIT.txt"],
  ["PyInstaller bootloader (6.15.0)", "LICENSES/PYINSTALLER-GPL2-EXCEPTION.txt"],
  ["Python runtime (3.11.9)", "LICENSES/PYTHON-3.11-PSF.txt"],
  ["SheetJS Community Edition (xlsx 0.20.3)", "node_modules/xlsx/LICENSE"],
  ["ExcelJS (@cweijan/exceljs 5.0.1)", "node_modules/@cweijan/exceljs/LICENSE"],
  ["React (react 19.2.8)", "node_modules/react/LICENSE"],
  ["React DOM (react-dom 19.2.8)", "node_modules/react-dom/LICENSE"],
  ["Ant Design (antd 5.29.3)", "node_modules/antd/LICENSE"],
  ["Ant Design Icons (@ant-design/icons 5.6.1)", "node_modules/@ant-design/icons/LICENSE"],
  ["VS Code Codicons (@vscode/codicons 0.0.45)", "node_modules/@vscode/codicons/LICENSE"],
  ["VS Code Codicons code license (@vscode/codicons 0.0.45)", "node_modules/@vscode/codicons/LICENSE-CODE"],
  ["JSZip (jszip 3.10.1)", "node_modules/jszip/LICENSE.markdown"],
  ["iconv-lite (iconv-lite 0.6.3)", "node_modules/iconv-lite/LICENSE"],
  ["uDSV (udsv 0.5.3)", "node_modules/udsv/LICENSE"],
  ["buffer (buffer 6.0.3)", "node_modules/buffer/LICENSE"],
];

const divider = "=".repeat(80);
const sections = [
  "THIRD-PARTY LICENSES",
  "",
  "The following license texts cover third-party components bundled in the",
  "Excel AI & VBA Studio VSIX. See THIRD_PARTY_NOTICES.md for source links.",
];

for (const [name, relativePath] of components) {
  const text = normalizeLineEndings(
    await readFile(path.join(root, relativePath), "utf8"),
  )
    .replace(/[ \t]+$/gm, "");
  sections.push("", divider, name, divider, "", text.trimEnd());
}

const generated = `${sections.join("\n")}\n`;

if (process.argv.includes("--check")) {
  let existing = "";
  try {
    existing = await readFile(outputPath, "utf8");
  } catch {
    // A missing output is reported as stale below.
  }

  if (normalizeLineEndings(existing) !== generated) {
    console.error("THIRD_PARTY_LICENSES.txt is missing or out of date.");
    process.exitCode = 1;
  } else {
    console.log("Third-party license bundle is up to date.");
  }
} else {
  await writeFile(outputPath, generated, "utf8");
  console.log(`Wrote ${path.relative(root, outputPath)}.`);
}
