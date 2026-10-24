import { createPortableFile } from './portableFileSave';
import { extname } from 'path';
import JSZip from 'jszip';
import * as vscode from 'vscode';

/** A minimal macro-free workbook; creation never starts Excel or reads user data. */
export async function createBlankWorkbookBytes(): Promise<Uint8Array> {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

export async function newWorkbook(): Promise<void> {
    const target = await vscode.window.showSaveDialog({
        title: 'Créer un classeur Excel',
        saveLabel: 'Créer le classeur',
        filters: { 'Classeur Excel': ['xlsx'] },
    });
    if (!target) return;
    if (target.scheme !== 'file' || extname(target.fsPath).toLowerCase() !== '.xlsx') {
        throw new Error('Choisissez un nouveau fichier local avec l’extension .xlsx.');
    }
    const bytes = await createBlankWorkbookBytes();
    // Exclusive creation also prevents replacing a symlink or an existing workbook.
    await createPortableFile(target.fsPath, bytes);
    await vscode.commands.executeCommand('vscode.openWith', target, 'excelAiVbaStudio.officeViewer');
}
