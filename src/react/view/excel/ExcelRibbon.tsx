import {
    AlignCenterOutlined,
    AlignLeftOutlined,
    AlignRightOutlined,
    ApartmentOutlined,
    BarChartOutlined,
    BgColorsOutlined,
    BoldOutlined,
    BorderOutlined,
    CalculatorOutlined,
    ClearOutlined,
    CodeOutlined,
    ColumnHeightOutlined,
    CommentOutlined,
    CopyOutlined,
    DeleteOutlined,
    DollarOutlined,
    DownOutlined,
    EyeOutlined,
    FileAddOutlined,
    FileExcelOutlined,
    FileTextOutlined,
    FilterOutlined,
    FontColorsOutlined,
    FormatPainterOutlined,
    FunctionOutlined,
    InsertRowAboveOutlined,
    InsertRowLeftOutlined,
    ItalicOutlined,
    LayoutOutlined,
    LinkOutlined,
    LockOutlined,
    MergeCellsOutlined,
    NumberOutlined,
    PercentageOutlined,
    PictureOutlined,
    PrinterOutlined,
    RedoOutlined,
    RobotOutlined,
    SaveOutlined,
    ScissorOutlined,
    SearchOutlined,
    SettingOutlined,
    SortAscendingOutlined,
    TableOutlined,
    UnderlineOutlined,
    UndoOutlined,
    UnlockOutlined,
    VerticalAlignMiddleOutlined,
    WindowsOutlined,
} from '@ant-design/icons';
import { Button, Input, Modal, Select } from 'antd';
import { useEffect, useState, type ReactNode } from 'react';
import type { SheetChartData, SheetTableData } from '../../../common/excelWorkbookObjects';
import type Spreadsheet from './x-spreadsheet/index';
import type { SheetConditionalFormattingRule } from './x-spreadsheet/index';
import ChartDesigner from './chart-designer';
import TableDesigner, { type TableDesignerValues } from './table-designer';
import './ExcelRibbon.less';

type RibbonTab =
    | 'file'
    | 'home'
    | 'insert'
    | 'layout'
    | 'formulas'
    | 'data'
    | 'review'
    | 'view'
    | 'ai-vba';

type ConditionalPreset =
    | 'greaterThan'
    | 'lessThan'
    | 'equal'
    | 'containsText'
    | 'colorScale'
    | 'dataBar'
    | 'iconSet';

type RibbonButtonProps = {
    icon: ReactNode;
    label: string;
    onClick?: () => void;
    disabled?: boolean;
    unavailable?: boolean;
    large?: boolean;
    compact?: boolean;
    title?: string;
    native?: boolean;
};

type ExcelRibbonProps = {
    spreadsheet: Spreadsheet | null;
    readOnly: boolean;
    allowSaveAs: boolean;
    showEditInVscode: boolean;
    onAutoFitColumns: () => void;
    onOpenExcel: () => void;
    onOpenVbe: () => void;
    onOpenVbaDeveloper: () => void;
    onExportWorkbookContext: () => void;
    onOpenVbaExplorer: () => void;
    onAskCopilotAboutWorkbook: (request?: string) => void;
};

type ChartSpreadsheet = Spreadsheet & {
    addOrUpdateChart: (chart: SheetChartData) => Spreadsheet;
    getActiveSheetCharts: () => SheetChartData[];
    getActiveSheetUnsupportedNativeChartCount: () => number;
    getSelectionRangeRef: () => string;
    removeChart: (chartId: string) => Spreadsheet;
};

type TableSpreadsheet = Spreadsheet & {
    formatSelectionAsTable: (options: TableDesignerValues) => Spreadsheet;
    getActiveSheetTables?: () => SheetTableData[];
    isWorkbookTableNameAvailable?: (name: string, currentTableId?: string) => boolean;
    getSelectionRangeRef?: () => string;
    removeTable?: (tableId: string) => Spreadsheet;
    updateTable?: (table: SheetTableData) => Spreadsheet;
};

const INTEGRATED_FEATURE_EVENT = 'excel-integrated-feature';

function chooseLocalFile(accept: string): Promise<File | null> {
    return new Promise(resolve => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = accept;
        input.style.display = 'none';
        input.addEventListener('change', () => {
            const file = input.files?.[0] ?? null;
            input.remove();
            resolve(file);
        }, { once: true });
        input.addEventListener('cancel', () => {
            input.remove();
            resolve(null);
        }, { once: true });
        document.body.appendChild(input);
        input.click();
    });
}

async function imagePayload(file: File): Promise<{
    base64: string;
    extension: 'jpeg' | 'png' | 'gif';
}> {
    const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result ?? ''));
        reader.onerror = () => reject(reader.error ?? new Error('Image illisible'));
        reader.readAsDataURL(file);
    });
    const match = /^data:image\/(png|jpe?g|gif);base64,(.+)$/i.exec(dataUrl);
    if (!match) throw new Error('Format image non pris en charge');
    return {
        extension: match[1].toLowerCase().startsWith('jp') ? 'jpeg' : match[1].toLowerCase() as 'png' | 'gif',
        base64: match[2],
    };
}

function generatedIllustration(feature: 'Shapes' | 'SmartArt', matrix: string[][]): string {
    const canvas = document.createElement('canvas');
    canvas.width = 720;
    canvas.height = 420;
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Canvas indisponible');
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.fillStyle = '#17365d';
    context.font = '700 24px Segoe UI';
    context.fillText(feature, 28, 42);

    if (feature.includes('Shape')) {
        context.fillStyle = '#5b9bd5';
        context.strokeStyle = '#2f5597';
        context.lineWidth = 4;
        context.beginPath();
        context.roundRect(150, 120, 420, 180, 24);
        context.fill();
        context.stroke();
        context.fillStyle = '#ffffff';
        context.textAlign = 'center';
        context.font = '700 30px Segoe UI';
        context.fillText('Forme Excel', 360, 220);
    } else {
        const labels = matrix.flat().filter(Boolean).slice(0, 3);
        ['Étape 1', 'Étape 2', 'Étape 3'].forEach((fallback, index) => {
            const x = 50 + index * 225;
            context.fillStyle = ['#4472c4', '#70ad47', '#ed7d31'][index];
            context.fillRect(x, 145, 170, 100);
            context.fillStyle = '#ffffff';
            context.textAlign = 'center';
            context.font = '600 18px Segoe UI';
            context.fillText(labels[index] || fallback, x + 85, 202, 150);
            if (index < 2) {
                context.fillStyle = '#7f8c8d';
                context.beginPath();
                context.moveTo(x + 180, 195);
                context.lineTo(x + 215, 175);
                context.lineTo(x + 215, 215);
                context.closePath();
                context.fill();
            }
        });
    }
    return canvas.toDataURL('image/png').split(',')[1];
}

const tabs: { id: RibbonTab; label: string }[] = [
    { id: 'file', label: 'File' },
    { id: 'home', label: 'Home' },
    { id: 'insert', label: 'Insert' },
    { id: 'layout', label: 'Page Layout' },
    { id: 'formulas', label: 'Formulas' },
    { id: 'data', label: 'Data' },
    { id: 'review', label: 'Review' },
    { id: 'view', label: 'View' },
    { id: 'ai-vba', label: 'AI & VBA' },
];

function RibbonButton({
    icon,
    label,
    onClick,
    disabled = false,
    unavailable = false,
    large = false,
    compact = false,
    title,
    native = false,
}: RibbonButtonProps) {
    const isDisabled = disabled || unavailable;
    const tooltip = title
        ?? (unavailable
            ? `${label} — not yet available in the embedded editor`
            : native
                ? `${label} — outil intégré dans VS Code`
                : label);

    return (
        <button
            type="button"
            className={[
                'excel-ribbon-button',
                large ? 'is-large' : '',
                compact ? 'is-compact' : '',
                unavailable ? 'is-unavailable' : '',
                native ? 'is-native' : '',
            ].filter(Boolean).join(' ')}
            onClick={native
                ? () => window.dispatchEvent(new CustomEvent(INTEGRATED_FEATURE_EVENT, { detail: label }))
                : onClick}
            disabled={isDisabled}
            title={tooltip}
        >
            <span className="excel-ribbon-button-icon" aria-hidden>{icon}</span>
            <span className="excel-ribbon-button-label">{label}</span>
        </button>
    );
}

function QuickButton({
    icon,
    label,
    onClick,
    disabled,
}: Pick<RibbonButtonProps, 'icon' | 'label' | 'onClick' | 'disabled'>) {
    return (
        <button
            type="button"
            className="excel-ribbon-quick-button"
            onClick={onClick}
            disabled={disabled}
            title={label}
            aria-label={label}
        >
            {icon}
        </button>
    );
}

function RibbonGroup({ label, children, wide = false }: { label: string; children: ReactNode; wide?: boolean }) {
    return (
        <section className={`excel-ribbon-group${wide ? ' is-wide' : ''}`} aria-label={label}>
            <div className="excel-ribbon-group-content">{children}</div>
            <div className="excel-ribbon-group-label">{label}</div>
        </section>
    );
}

function RibbonStack({ children }: { children: ReactNode }) {
    return <div className="excel-ribbon-stack">{children}</div>;
}

function RibbonSelect({
    label,
    defaultValue,
    onChange,
    disabled,
    children,
    wide = false,
}: {
    label: string;
    defaultValue: string;
    onChange: (value: string) => void;
    disabled?: boolean;
    children: ReactNode;
    wide?: boolean;
}) {
    return (
        <label className={`excel-ribbon-select${wide ? ' is-wide' : ''}`} title={label}>
            <span className="sr-only">{label}</span>
            <select
                aria-label={label}
                defaultValue={defaultValue}
                disabled={disabled}
                onChange={(event) => onChange(event.target.value)}
            >
                {children}
            </select>
            <DownOutlined aria-hidden />
        </label>
    );
}

function ColorButton({
    label,
    icon,
    defaultValue,
    disabled,
    onChange,
}: {
    label: string;
    icon: ReactNode;
    defaultValue: string;
    disabled?: boolean;
    onChange: (value: string) => void;
}) {
    return (
        <label className={`excel-ribbon-color${disabled ? ' is-disabled' : ''}`} title={label}>
            <span className="excel-ribbon-color-icon" aria-hidden>{icon}</span>
            <span className="excel-ribbon-color-line" style={{ backgroundColor: defaultValue }} />
            <input
                type="color"
                aria-label={label}
                defaultValue={defaultValue}
                disabled={disabled}
                onChange={(event) => onChange(event.target.value)}
            />
        </label>
    );
}

export default function ExcelRibbon({
    spreadsheet,
    readOnly,
    allowSaveAs,
    showEditInVscode,
    onAutoFitColumns,
    onOpenExcel,
    onOpenVbe,
    onOpenVbaDeveloper,
    onExportWorkbookContext,
    onOpenVbaExplorer,
    onAskCopilotAboutWorkbook,
}: ExcelRibbonProps) {
    const [activeTab, setActiveTab] = useState<RibbonTab>('home');
    const [gridVisible, setGridVisibleState] = useState(true);
    const [sheetProtected, setSheetProtected] = useState(false);
    const [conditionalOpen, setConditionalOpen] = useState(false);
    const [conditionalPreset, setConditionalPreset] = useState<ConditionalPreset>('greaterThan');
    const [conditionalValue, setConditionalValue] = useState('0');
    const [integratedFeature, setIntegratedFeature] = useState<string | null>(null);
    const [featureRunning, setFeatureRunning] = useState(false);
    const [featureResult, setFeatureResult] = useState('');
    const [headingsVisible, setHeadingsVisible] = useState(true);
    const [formulaBarVisible, setFormulaBarVisible] = useState(true);
    const [chartDesignerOpen, setChartDesignerOpen] = useState(false);
    const [chartInitialType, setChartInitialType] = useState(51);
    const [, setChartRevision] = useState(0);
    const [tableDesignerOpen, setTableDesignerOpen] = useState(false);
    const [, setTableRevision] = useState(0);
    const ready = spreadsheet !== null;
    const editingDisabled = !ready || readOnly;
    const chartSpreadsheet = spreadsheet as ChartSpreadsheet | null;
    const tableSpreadsheet = spreadsheet as TableSpreadsheet | null;
    const activeCharts = chartSpreadsheet?.getActiveSheetCharts?.() ?? [];
    const unsupportedNativeChartCount = chartSpreadsheet?.getActiveSheetUnsupportedNativeChartCount?.() ?? 0;
    const activeTables = tableSpreadsheet?.getActiveSheetTables?.() ?? [];
    const tableInventoryAvailable = typeof tableSpreadsheet?.getActiveSheetTables === 'function';
    const selectionRangeRef = chartSpreadsheet?.getSelectionRangeRef?.()
        ?? tableSpreadsheet?.getSelectionRangeRef?.()
        ?? '';

    useEffect(() => {
        setGridVisibleState(true);
        setSheetProtected(spreadsheet?.isSheetProtected() ?? false);
    }, [spreadsheet]);

    useEffect(() => {
        const onIntegratedFeature = (event: Event) => {
            const feature = (event as CustomEvent<string>).detail;
            if (!feature) return;
            setFeatureResult('');
            setIntegratedFeature(feature);
        };
        window.addEventListener(INTEGRATED_FEATURE_EVENT, onIntegratedFeature);
        return () => window.removeEventListener(INTEGRATED_FEATURE_EVENT, onIntegratedFeature);
    }, []);

    const command = (type: string, value?: unknown) => {
        spreadsheet?.executeCommand(type, value);
    };
    const toggle = (type: string) => {
        spreadsheet?.toggleCommand(type);
    };
    const context = (type: string) => {
        spreadsheet?.executeContextCommand(type);
    };
    const toggleGrid = () => {
        const next = !gridVisible;
        setGridVisibleState(next);
        spreadsheet?.setGridVisible(next);
    };
    const openNativeFeature = () => {};
    const openChartDesigner = (chartType: number) => {
        if (!chartSpreadsheet) return;
        setChartInitialType(chartType);
        setChartDesignerOpen(true);
    };
    const saveChart = (chart: SheetChartData) => {
        if (!chartSpreadsheet) throw new Error('Aucune feuille active pour enregistrer le graphique.');
        chartSpreadsheet.addOrUpdateChart(chart);
        setChartRevision(revision => revision + 1);
    };
    const removeChart = (chartId: string) => {
        if (!chartSpreadsheet) return;
        chartSpreadsheet.removeChart(chartId);
        setChartRevision(revision => revision + 1);
    };
    const createTable = (values: TableDesignerValues) => {
        if (!tableSpreadsheet) throw new Error('Aucune feuille active pour créer la table.');
        tableSpreadsheet.formatSelectionAsTable(values);
        setTableRevision(revision => revision + 1);
    };
    const updateTable = tableSpreadsheet?.updateTable
        ? (table: SheetTableData) => {
            tableSpreadsheet.updateTable?.(table);
            setTableRevision(revision => revision + 1);
        }
        : undefined;
    const removeTable = tableSpreadsheet?.removeTable
        ? (tableId: string) => {
            tableSpreadsheet.removeTable?.(tableId);
            setTableRevision(revision => revision + 1);
        }
        : undefined;

    const importLocalWorkbook = async (file: File) => {
        if (!spreadsheet) return;
        const extension = file.name.split('.').pop()?.toLowerCase() || 'xlsx';
        const { loadSheets } = await import('./excel_reader.ts');
        const imported = await loadSheets(await file.arrayBuffer(), extension);
        spreadsheet.appendSheets(imported.sheets);
    };

    const runIntegratedFeature = async () => {
        if (!spreadsheet || !integratedFeature) return;
        const feature = integratedFeature;
        setFeatureRunning(true);
        setFeatureResult('');
        try {
            switch (feature) {
                case 'New Workbook':
                    spreadsheet.addBlankSheet();
                    setFeatureResult('Une nouvelle feuille vierge a été ajoutée au classeur.');
                    break;
                case 'Options': {
                    const zoom = window.prompt('Zoom du classeur (50 à 200 %)', '100');
                    if (zoom != null) spreadsheet.setZoom(Math.max(0.5, Math.min(2, Number(zoom) / 100 || 1)));
                    setFeatureResult('Les options d’affichage ont été appliquées.');
                    break;
                }
                case 'PivotTable':
                    spreadsheet.copySelectionToNewSheet('Tableau croisé').addSubtotal();
                    setFeatureResult('Une synthèse de la sélection a été créée dans une nouvelle feuille.');
                    break;
                case 'Pictures':
                case 'Background': {
                    const file = await chooseLocalFile('image/png,image/jpeg,image/gif');
                    if (!file) break;
                    const payload = await imagePayload(file);
                    spreadsheet.insertImage(payload.base64, payload.extension, feature === 'Background');
                    setFeatureResult(feature === 'Background'
                        ? 'L’image a été appliquée comme arrière-plan.'
                        : 'L’image a été insérée sur la feuille.');
                    break;
                }
                case 'Shapes':
                case 'SmartArt': {
                    const base64 = generatedIllustration(feature, spreadsheet.getSelectionMatrix());
                    spreadsheet.insertImage(base64, 'png');
                    setFeatureResult(`${feature} a été généré dans la feuille et sera conservé dans le fichier XLSX.`);
                    break;
                }
                case 'Themes':
                    command('font-name', 'Segoe UI');
                    command('color', '#1f1f1f');
                    command('bgcolor', '#ffffff');
                    setFeatureResult('Le thème Office moderne a été appliqué à la sélection.');
                    break;
                case 'Colors':
                    command('bgcolor', '#ddebf7');
                    command('color', '#17365d');
                    setFeatureResult('La palette Office bleue a été appliquée.');
                    break;
                case 'Fonts':
                    command('font-name', 'Segoe UI');
                    setFeatureResult('La police Segoe UI a été appliquée à la sélection.');
                    break;
                case 'Effects':
                    command('border', { mode: 'all', style: 'thin', color: '#9eafbf' });
                    setFeatureResult('Les bordures et effets de tableau ont été appliqués.');
                    break;
                case 'Margins':
                    spreadsheet.setPageSetup({
                        margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
                    });
                    setFeatureResult('Les marges normales ont été enregistrées.');
                    break;
                case 'Orientation':
                    spreadsheet.setPageSetup({ orientation: 'landscape' });
                    setFeatureResult('La feuille est configurée en orientation paysage.');
                    break;
                case 'Size':
                    spreadsheet.setPageSetup({ paperSize: 9 });
                    setFeatureResult('Le format A4 a été enregistré.');
                    break;
                case 'Headings': {
                    const next = !headingsVisible;
                    setHeadingsVisible(next);
                    spreadsheet.setHeadingsVisible(next);
                    setFeatureResult(next ? 'Les en-têtes sont visibles.' : 'Les en-têtes sont masqués.');
                    break;
                }
                case 'Width: Auto':
                    spreadsheet.autoFitColumns();
                    setFeatureResult('La largeur des colonnes a été ajustée.');
                    break;
                case 'Height: Auto':
                    spreadsheet.autoFitRows();
                    setFeatureResult('La hauteur des lignes a été ajustée.');
                    break;
                case 'Bring Forward':
                    spreadsheet.arrangeSelectedImage('forward');
                    setFeatureResult('L’objet sélectionné a été avancé.');
                    break;
                case 'Send Backward':
                    spreadsheet.arrangeSelectedImage('backward');
                    setFeatureResult('L’objet sélectionné a été reculé.');
                    break;
                case 'Selection Pane': {
                    const sheet = spreadsheet.getData()[spreadsheet.getActiveSheetIndex()];
                    const count = sheet?.images?.length ?? 0;
                    setFeatureResult(`${count} objet${count === 1 ? '' : 's'} graphique${count === 1 ? '' : 's'} dans la feuille active.`);
                    break;
                }
                case 'Trace Precedents': {
                    const items = spreadsheet.formulaAudit('precedents');
                    setFeatureResult(items.length ? `Antécédents : ${items.join(', ')}` : 'Aucun antécédent direct détecté.');
                    break;
                }
                case 'Trace Dependents': {
                    const items = spreadsheet.formulaAudit('dependents');
                    setFeatureResult(items.length ? `Dépendants : ${items.join(', ')}` : 'Aucun dépendant direct détecté.');
                    break;
                }
                case 'Show Formulas': {
                    const showing = spreadsheet.toggleFormulaDisplay();
                    setFeatureResult(showing ? 'Les formules sont affichées.' : 'Les résultats sont affichés.');
                    break;
                }
                case 'Error Checking': {
                    const items = spreadsheet.formulaAudit('errors');
                    setFeatureResult(items.length ? `Erreurs détectées : ${items.join(', ')}` : 'Aucune erreur de formule détectée.');
                    break;
                }
                case 'Calculate Now':
                case 'Calculation Options':
                    spreadsheet.reRender();
                    setFeatureResult('Le classeur a été recalculé et actualisé.');
                    break;
                case 'Get Data':
                case 'From Text/CSV': {
                    const file = await chooseLocalFile('.xlsx,.xls,.ods,.csv,.tsv');
                    if (!file) break;
                    await importLocalWorkbook(file);
                    setFeatureResult(`Les données de ${file.name} ont été ajoutées dans une nouvelle feuille.`);
                    break;
                }
                case 'From Web': {
                    const url = window.prompt('Adresse HTTPS du fichier CSV, TSV, XLSX ou XLS à importer');
                    if (!url) break;
                    const response = await fetch(url);
                    if (!response.ok) throw new Error(`Téléchargement impossible (${response.status})`);
                    const extension = new URL(url).pathname.split('.').pop()?.toLowerCase() || 'csv';
                    const { loadSheets } = await import('./excel_reader.ts');
                    const imported = await loadSheets(await response.arrayBuffer(), extension);
                    spreadsheet.appendSheets(imported.sheets);
                    setFeatureResult('Les données Web ont été ajoutées au classeur.');
                    break;
                }
                case 'From Table/Range':
                    spreadsheet.copySelectionToNewSheet();
                    setFeatureResult('La sélection a été copiée dans une nouvelle feuille.');
                    break;
                case 'Text to Columns': {
                    const delimiter = window.prompt('Séparateur à utiliser', ',');
                    if (delimiter == null) break;
                    spreadsheet.textToColumns(delimiter);
                    setFeatureResult('Le texte a été réparti en colonnes.');
                    break;
                }
                case 'Remove Duplicates': {
                    const removed = spreadsheet.removeDuplicateRows();
                    setFeatureResult(`${removed} ligne${removed === 1 ? '' : 's'} en double supprimée${removed === 1 ? '' : 's'}.`);
                    break;
                }
                case 'Subtotal':
                    spreadsheet.addSubtotal();
                    setFeatureResult('Une ligne de sous-total a été ajoutée.');
                    break;
                case 'What-If Analysis': {
                    const value = window.prompt('Nouvelle valeur pour la cellule sélectionnée');
                    if (value == null) break;
                    const selection = spreadsheet.getSelection();
                    spreadsheet.cellText(selection.ri, selection.ci, value);
                    setFeatureResult('La valeur du scénario a été appliquée.');
                    break;
                }
                case 'Forecast Sheet':
                    spreadsheet.addForecastRow();
                    setFeatureResult('Une ligne de prévision linéaire a été ajoutée sous la sélection.');
                    break;
                case 'Protect Workbook': {
                    const active = spreadsheet.toggleWorkbookProtection();
                    setFeatureResult(active ? 'Toutes les feuilles sont protégées.' : 'La protection du classeur est retirée.');
                    break;
                }
                case 'Page Break Preview':
                case 'Page Layout':
                    command('print');
                    setFeatureResult('L’aperçu de mise en page est ouvert dans VS Code.');
                    break;
                case 'Formula Bar': {
                    const next = !formulaBarVisible;
                    setFormulaBarVisible(next);
                    spreadsheet.setFormulaBarVisible(next);
                    setFeatureResult(next ? 'La barre de formule est visible.' : 'La barre de formule est masquée.');
                    break;
                }
                case 'Zoom': {
                    const value = window.prompt('Zoom (50 à 200 %)', '100');
                    if (value != null) spreadsheet.setZoom(Math.max(0.5, Math.min(2, Number(value) / 100 || 1)));
                    setFeatureResult('Le niveau de zoom a été appliqué.');
                    break;
                }
                case 'Zoom to Selection':
                    spreadsheet.setZoom(1.35);
                    setFeatureResult('La sélection a été agrandie.');
                    break;
                case 'Split':
                    spreadsheet.toggleCommand('freeze');
                    setFeatureResult('Le fractionnement a été appliqué à la sélection.');
                    break;
                default:
                    onAskCopilotAboutWorkbook(
                        `Dans VS Code uniquement, réalise l’action Excel « ${feature} » sur le classeur actif. `
                        + 'Analyse les données et le VBA avec #excelVbaWorkbook, puis propose ou applique la solution sans ouvrir Microsoft Excel.',
                    );
                    setFeatureResult(`La demande « ${feature} » a été transmise à GitHub Copilot avec le contexte du classeur.`);
                    break;
            }
        } catch (error) {
            setFeatureResult(`Erreur : ${(error as Error).message}`);
        } finally {
            setFeatureRunning(false);
        }
    };
    const applyConditionalFormatting = () => {
        if (!spreadsheet) return;
        const numericValue = Number(conditionalValue.replace(',', '.'));
        const formulaValue = Number.isNaN(numericValue) ? conditionalValue : numericValue;
        let rule: SheetConditionalFormattingRule;
        if (conditionalPreset === 'colorScale') {
            rule = {
                type: 'colorScale',
                cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
                color: [
                    { argb: 'FFF8696B' },
                    { argb: 'FFFFEB84' },
                    { argb: 'FF63BE7B' },
                ],
            };
        } else if (conditionalPreset === 'dataBar') {
            rule = {
                type: 'dataBar',
                cfvo: [{ type: 'min' }, { type: 'max' }],
                color: { argb: 'FF5B9BD5' },
            };
        } else if (conditionalPreset === 'iconSet') {
            rule = {
                type: 'iconSet',
                iconSet: '3TrafficLights1',
                cfvo: [
                    { type: 'min' },
                    { type: 'percent', value: 33 },
                    { type: 'percent', value: 67 },
                ],
            };
        } else {
            const isContains = conditionalPreset === 'containsText';
            rule = {
                type: isContains ? 'containsText' : 'cellIs',
                ...(isContains
                    ? { operator: 'containsText', text: conditionalValue, formulae: [conditionalValue] }
                    : { operator: conditionalPreset, formulae: [formulaValue] }),
                style: {
                    fill: {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFC7CE' },
                    },
                    font: { color: { argb: 'FF9C0006' } },
                },
                displayStyle: {
                    bgcolor: '#ffc7ce',
                    color: '#9c0006',
                    font: { bold: true },
                },
            };
        }
        spreadsheet.addConditionalFormatting(rule);
        setConditionalOpen(false);
    };
    const addComment = () => {
        if (!spreadsheet) return;
        const current = spreadsheet.getSelectedComment()?.text ?? '';
        const text = window.prompt('Comment for the selected cell (empty removes it)', current);
        if (text == null) return;
        spreadsheet.setSelectedComment(text);
    };
    const showComments = () => {
        const comments = spreadsheet?.listComments() ?? [];
        window.alert(comments.length
            ? comments.map(comment => (
                `${comment.sheet}!${comment.address}: ${comment.text}`
            )).join('\n\n')
            : 'This workbook has no comments.');
    };
    const showWorkbookStatistics = () => {
        const statistics = spreadsheet?.getWorkbookStatistics();
        if (!statistics) return;
        window.alert([
            `Sheets: ${statistics.sheets}`,
            `Populated cells: ${statistics.populatedCells}`,
            `Formulas: ${statistics.formulas}`,
            `Comments: ${statistics.comments}`,
            `Conditional formatting rules: ${statistics.conditionalFormattingRules}`,
        ].join('\n'));
    };
    const toggleSheetProtection = () => {
        if (!spreadsheet) return;
        setSheetProtected(spreadsheet.toggleSheetProtection());
    };

    const renderFileTab = () => (
        <>
            <RibbonGroup label="Workbook">
                <RibbonButton large icon={<SaveOutlined />} label="Save" onClick={() => command('save')} disabled={editingDisabled} />
                <RibbonButton large icon={<FileAddOutlined />} label="Save As" onClick={() => command('save-as')} disabled={!ready || !allowSaveAs} />
                <RibbonButton large icon={<PrinterOutlined />} label="Print" onClick={() => command('print')} disabled={!ready} />
            </RibbonGroup>
            <RibbonGroup label="Open">
                <RibbonButton large icon={<FileExcelOutlined />} label="Open in Excel" onClick={onOpenExcel} />
                <RibbonButton large icon={<WindowsOutlined />} label="Open native VBE" onClick={onOpenVbe} />
                <RibbonButton large icon={<CodeOutlined />} label="VBA Studio" onClick={onOpenVbaDeveloper} />
                {showEditInVscode && (
                    <RibbonButton large icon={<FileTextOutlined />} label="Edit as Text" onClick={() => command('edit-in-vscode')} disabled={!ready} />
                )}
            </RibbonGroup>
            <RibbonGroup label="Create">
                <RibbonButton large icon={<FileAddOutlined />} label="New Workbook" onClick={openNativeFeature} native />
                <RibbonButton large icon={<SettingOutlined />} label="Options" onClick={openNativeFeature} native />
            </RibbonGroup>
        </>
    );

    const renderHomeTab = () => (
        <>
            <RibbonGroup label="Clipboard">
                <RibbonButton large icon={<CopyOutlined />} label="Paste" onClick={() => context('paste')} disabled={editingDisabled} />
                <RibbonStack>
                    <RibbonButton compact icon={<ScissorOutlined />} label="Cut" onClick={() => context('cut')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<CopyOutlined />} label="Copy" onClick={() => context('copy')} disabled={!ready} />
                    <RibbonButton compact icon={<FormatPainterOutlined />} label="Format Painter" onClick={() => toggle('paintformat')} disabled={editingDisabled} />
                </RibbonStack>
            </RibbonGroup>
            <RibbonGroup label="Font" wide>
                <div className="excel-ribbon-control-row">
                    <RibbonSelect label="Font" defaultValue="Calibri" onChange={(value) => command('font-name', value)} disabled={editingDisabled} wide>
                        {['Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Georgia', 'Courier New'].map((font) => (
                            <option key={font} value={font}>{font}</option>
                        ))}
                    </RibbonSelect>
                    <RibbonSelect label="Font size" defaultValue="11" onChange={(value) => command('font-size', Number(value))} disabled={editingDisabled}>
                        {[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72].map((size) => (
                            <option key={size} value={size}>{size}</option>
                        ))}
                    </RibbonSelect>
                </div>
                <div className="excel-ribbon-control-row">
                    <RibbonButton compact icon={<BoldOutlined />} label="Bold" onClick={() => toggle('font-bold')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<ItalicOutlined />} label="Italic" onClick={() => toggle('font-italic')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<UnderlineOutlined />} label="Underline" onClick={() => toggle('underline')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<BorderOutlined />} label="Borders" onClick={() => command('border', { mode: 'all', style: 'thin', color: '#7f8c8d' })} disabled={editingDisabled} />
                    <ColorButton label="Fill color" icon={<BgColorsOutlined />} defaultValue="#fff2cc" disabled={editingDisabled} onChange={(value) => command('bgcolor', value)} />
                    <ColorButton label="Font color" icon={<FontColorsOutlined />} defaultValue="#c00000" disabled={editingDisabled} onChange={(value) => command('color', value)} />
                </div>
            </RibbonGroup>
            <RibbonGroup label="Alignment" wide>
                <div className="excel-ribbon-control-row">
                    <RibbonButton compact icon={<AlignLeftOutlined />} label="Left" onClick={() => command('align', 'left')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<AlignCenterOutlined />} label="Center" onClick={() => command('align', 'center')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<AlignRightOutlined />} label="Right" onClick={() => command('align', 'right')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<VerticalAlignMiddleOutlined />} label="Middle" onClick={() => command('valign', 'middle')} disabled={editingDisabled} />
                </div>
                <div className="excel-ribbon-control-row">
                    <RibbonButton compact icon={<ColumnHeightOutlined />} label="Wrap Text" onClick={() => toggle('textwrap')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<MergeCellsOutlined />} label="Merge & Center" onClick={() => toggle('merge')} disabled={editingDisabled} />
                </div>
            </RibbonGroup>
            <RibbonGroup label="Number">
                <RibbonSelect label="Number format" defaultValue="normal" onChange={(value) => command('format', value)} disabled={editingDisabled} wide>
                    <option value="normal">General</option>
                    <option value="number">Number</option>
                    <option value="usd">Currency ($)</option>
                    <option value="eur">Currency (€)</option>
                    <option value="percent">Percentage</option>
                    <option value="date">Date</option>
                    <option value="time">Time</option>
                </RibbonSelect>
                <div className="excel-ribbon-control-row">
                    <RibbonButton compact icon={<DollarOutlined />} label="Currency" onClick={() => command('format', 'eur')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<PercentageOutlined />} label="Percent" onClick={() => command('format', 'percent')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<NumberOutlined />} label="Number" onClick={() => command('format', 'number')} disabled={editingDisabled} />
                </div>
            </RibbonGroup>
            <RibbonGroup label="Styles">
                <RibbonButton icon={<ClearOutlined />} label="Clear Formatting" onClick={() => command('clearformat')} disabled={editingDisabled} />
                <RibbonButton icon={<TableOutlined />} label="Format as Table" onClick={() => setTableDesignerOpen(true)} disabled={!ready} />
                <RibbonButton icon={<BarChartOutlined />} label="Conditional Formatting" onClick={() => setConditionalOpen(true)} disabled={editingDisabled} />
            </RibbonGroup>
            <RibbonGroup label="Cells">
                <RibbonButton compact icon={<InsertRowAboveOutlined />} label="Insert Row" onClick={() => context('insert-row')} disabled={editingDisabled} />
                <RibbonButton compact icon={<InsertRowLeftOutlined />} label="Insert Column" onClick={() => context('insert-column')} disabled={editingDisabled} />
                <RibbonButton compact icon={<DeleteOutlined />} label="Delete" onClick={() => context('delete-cell-text')} disabled={editingDisabled} />
                <RibbonButton compact icon={<ColumnHeightOutlined />} label="AutoFit Columns" onClick={onAutoFitColumns} disabled={!ready} />
            </RibbonGroup>
            <RibbonGroup label="Editing">
                <RibbonButton large icon={<FunctionOutlined />} label="AutoSum" onClick={() => command('formula', 'SUM')} disabled={editingDisabled} />
                <RibbonStack>
                    <RibbonButton compact icon={<FilterOutlined />} label="Sort & Filter" onClick={() => toggle('autofilter')} disabled={editingDisabled} />
                    <RibbonButton compact icon={<SearchOutlined />} label="Find & Select" onClick={() => command('find')} disabled={!ready} />
                    <RibbonButton compact icon={<ClearOutlined />} label="Clear Contents" onClick={() => context('delete-cell-text')} disabled={editingDisabled} />
                </RibbonStack>
            </RibbonGroup>
        </>
    );

    const renderInsertTab = () => (
        <>
            <RibbonGroup label="Tables">
                <RibbonButton large icon={<TableOutlined />} label="PivotTable" onClick={openNativeFeature} native />
                <RibbonButton large icon={<TableOutlined />} label="Table" onClick={() => setTableDesignerOpen(true)} disabled={!ready} />
            </RibbonGroup>
            <RibbonGroup label="Cells">
                <RibbonButton large icon={<InsertRowAboveOutlined />} label="Insert Row" onClick={() => context('insert-row')} disabled={editingDisabled} />
                <RibbonButton large icon={<InsertRowLeftOutlined />} label="Insert Column" onClick={() => context('insert-column')} disabled={editingDisabled} />
                <RibbonButton large icon={<DeleteOutlined />} label="Delete Cells" onClick={() => context('delete-cell')} disabled={editingDisabled} />
            </RibbonGroup>
            <RibbonGroup label="Illustrations">
                <RibbonButton large icon={<PictureOutlined />} label="Pictures" onClick={openNativeFeature} native />
                <RibbonButton large icon={<LayoutOutlined />} label="Shapes" onClick={openNativeFeature} native />
                <RibbonButton large icon={<ApartmentOutlined />} label="SmartArt" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Charts">
				<RibbonButton large icon={<BarChartOutlined />} label="All Charts" onClick={() => openChartDesigner(51)} disabled={!ready} />
                <RibbonButton large icon={<BarChartOutlined />} label="Column Chart" onClick={() => openChartDesigner(51)} disabled={!ready} />
                <RibbonButton large icon={<BarChartOutlined />} label="Line Chart" onClick={() => openChartDesigner(4)} disabled={!ready} />
            </RibbonGroup>
            <RibbonGroup label="Links">
                <RibbonButton large icon={<LinkOutlined />} label="Link" onClick={() => context('hyperlink')} disabled={editingDisabled} />
            </RibbonGroup>
            <RibbonGroup label="Functions">
                <RibbonButton large icon={<FunctionOutlined />} label="Insert Function" onClick={() => command('formula', 'SUM')} disabled={editingDisabled} />
            </RibbonGroup>
        </>
    );

    const renderLayoutTab = () => (
        <>
            <RibbonGroup label="Themes">
                <RibbonButton large icon={<LayoutOutlined />} label="Themes" onClick={openNativeFeature} native />
                <RibbonStack>
                    <RibbonButton compact icon={<BgColorsOutlined />} label="Colors" onClick={openNativeFeature} native />
                    <RibbonButton compact icon={<FontColorsOutlined />} label="Fonts" onClick={openNativeFeature} native />
                    <RibbonButton compact icon={<SettingOutlined />} label="Effects" onClick={openNativeFeature} native />
                </RibbonStack>
            </RibbonGroup>
            <RibbonGroup label="Page Setup">
                <RibbonButton large icon={<LayoutOutlined />} label="Margins" onClick={openNativeFeature} native />
                <RibbonButton large icon={<LayoutOutlined />} label="Orientation" onClick={openNativeFeature} native />
                <RibbonButton large icon={<FileTextOutlined />} label="Size" onClick={openNativeFeature} native />
                <RibbonButton large icon={<PrinterOutlined />} label="Print Area" onClick={() => command('print')} disabled={!ready} />
                <RibbonButton large icon={<PictureOutlined />} label="Background" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Sheet Options">
                <RibbonButton large icon={<EyeOutlined />} label={gridVisible ? 'Hide Gridlines' : 'Show Gridlines'} onClick={toggleGrid} disabled={!ready} />
                <RibbonButton large icon={<ColumnHeightOutlined />} label="Headings" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Scale to Fit">
                <RibbonButton large icon={<ColumnHeightOutlined />} label="Width: Auto" onClick={openNativeFeature} native />
                <RibbonButton large icon={<ColumnHeightOutlined />} label="Height: Auto" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Arrange">
                <RibbonButton large icon={<LayoutOutlined />} label="Bring Forward" onClick={openNativeFeature} native />
                <RibbonButton large icon={<LayoutOutlined />} label="Send Backward" onClick={openNativeFeature} native />
                <RibbonButton large icon={<LayoutOutlined />} label="Selection Pane" onClick={openNativeFeature} native />
            </RibbonGroup>
        </>
    );

    const renderFormulasTab = () => (
        <>
            <RibbonGroup label="Function Library">
                <RibbonButton large icon={<FunctionOutlined />} label="Insert Function" onClick={() => command('formula', 'SUM')} disabled={editingDisabled} />
                <RibbonButton large icon={<CalculatorOutlined />} label="AutoSum" onClick={() => command('formula', 'SUM')} disabled={editingDisabled} />
                <RibbonButton large icon={<FunctionOutlined />} label="Average" onClick={() => command('formula', 'AVERAGE')} disabled={editingDisabled} />
                <RibbonButton large icon={<FunctionOutlined />} label="Maximum" onClick={() => command('formula', 'MAX')} disabled={editingDisabled} />
                <RibbonButton large icon={<FunctionOutlined />} label="Minimum" onClick={() => command('formula', 'MIN')} disabled={editingDisabled} />
                <RibbonButton large icon={<FunctionOutlined />} label="Logical IF" onClick={() => command('formula', 'IF')} disabled={editingDisabled} />
            </RibbonGroup>
            <RibbonGroup label="Defined Names">
                <RibbonButton large icon={<FileTextOutlined />} label="Name Manager" onClick={openNativeFeature} native />
                <RibbonButton large icon={<FileAddOutlined />} label="Define Name" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Formula Auditing">
                <RibbonButton large icon={<ApartmentOutlined />} label="Trace Precedents" onClick={openNativeFeature} native />
                <RibbonButton large icon={<ApartmentOutlined />} label="Trace Dependents" onClick={openNativeFeature} native />
                <RibbonButton large icon={<EyeOutlined />} label="Show Formulas" onClick={openNativeFeature} native />
                <RibbonButton large icon={<SearchOutlined />} label="Error Checking" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Calculation">
                <RibbonButton large icon={<CalculatorOutlined />} label="Calculate Now" onClick={openNativeFeature} native />
                <RibbonButton large icon={<SettingOutlined />} label="Calculation Options" onClick={openNativeFeature} native />
            </RibbonGroup>
        </>
    );

    const renderDataTab = () => (
        <>
            <RibbonGroup label="Get & Transform Data">
                <RibbonButton large icon={<FileTextOutlined />} label="Get Data" onClick={openNativeFeature} native />
                <RibbonButton large icon={<FileTextOutlined />} label="From Text/CSV" onClick={openNativeFeature} native />
                <RibbonButton large icon={<FileTextOutlined />} label="From Web" onClick={openNativeFeature} native />
                <RibbonButton large icon={<TableOutlined />} label="From Table/Range" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Sort & Filter">
                <RibbonButton large icon={<SortAscendingOutlined />} label="Sort A to Z" onClick={() => spreadsheet?.sortSelection('asc')} disabled={!ready} />
                <RibbonButton large icon={<FilterOutlined />} label="Filter" onClick={() => toggle('autofilter')} disabled={editingDisabled} />
                <RibbonButton large icon={<ClearOutlined />} label="Clear" onClick={() => context('delete-cell-text')} disabled={editingDisabled} />
            </RibbonGroup>
            <RibbonGroup label="Data Tools">
                <RibbonButton large icon={<ColumnHeightOutlined />} label="Text to Columns" onClick={openNativeFeature} native />
                <RibbonButton large icon={<DeleteOutlined />} label="Remove Duplicates" onClick={openNativeFeature} native />
                <RibbonButton large icon={<SettingOutlined />} label="Data Validation" onClick={() => context('validation')} disabled={editingDisabled} />
            </RibbonGroup>
            <RibbonGroup label="Outline">
                <RibbonButton large icon={<LayoutOutlined />} label="Group" onClick={openNativeFeature} native />
                <RibbonButton large icon={<LayoutOutlined />} label="Ungroup" onClick={openNativeFeature} native />
                <RibbonButton large icon={<CalculatorOutlined />} label="Subtotal" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Analyze">
                <RibbonButton large icon={<RobotOutlined />} label="What-If Analysis" onClick={openNativeFeature} native />
                <RibbonButton large icon={<BarChartOutlined />} label="Forecast Sheet" onClick={openNativeFeature} native />
            </RibbonGroup>
        </>
    );

    const renderReviewTab = () => (
        <>
            <RibbonGroup label="Proofing">
                <RibbonButton large icon={<SearchOutlined />} label="Spelling" onClick={openNativeFeature} native />
                <RibbonButton large icon={<FileTextOutlined />} label="Thesaurus" onClick={openNativeFeature} native />
                <RibbonButton large icon={<NumberOutlined />} label="Workbook Statistics" onClick={showWorkbookStatistics} disabled={!ready} />
            </RibbonGroup>
            <RibbonGroup label="Language">
                <RibbonButton large icon={<RobotOutlined />} label="Translate" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Comments">
                <RibbonButton large icon={<CommentOutlined />} label="New Comment" onClick={addComment} disabled={editingDisabled} />
                <RibbonButton large icon={<CommentOutlined />} label="Show Comments" onClick={showComments} disabled={!ready} />
            </RibbonGroup>
            <RibbonGroup label="Protect">
                <RibbonButton large icon={<LockOutlined />} label="Lock Cells" onClick={() => context('cell-non-editable')} disabled={editingDisabled} />
                <RibbonButton large icon={<UnlockOutlined />} label="Unlock Cells" onClick={() => context('cell-editable')} disabled={editingDisabled} />
                <RibbonButton
                    large
                    icon={sheetProtected ? <UnlockOutlined /> : <LockOutlined />}
                    label={sheetProtected ? 'Unprotect Sheet' : 'Protect Sheet'}
                    onClick={toggleSheetProtection}
                    disabled={editingDisabled}
                />
                <RibbonButton large icon={<LockOutlined />} label="Protect Workbook" onClick={openNativeFeature} native />
            </RibbonGroup>
        </>
    );

    const renderViewTab = () => (
        <>
            <RibbonGroup label="Workbook Views">
                <RibbonButton large icon={<TableOutlined />} label="Normal" onClick={() => spreadsheet?.setZoom(1)} disabled={!ready} />
                <RibbonButton large icon={<LayoutOutlined />} label="Page Break Preview" onClick={openNativeFeature} native />
                <RibbonButton large icon={<FileTextOutlined />} label="Page Layout" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Show">
                <RibbonButton large icon={<EyeOutlined />} label={gridVisible ? 'Gridlines On' : 'Gridlines Off'} onClick={toggleGrid} disabled={!ready} />
                <RibbonButton large icon={<ColumnHeightOutlined />} label="Headings" onClick={openNativeFeature} native />
                <RibbonButton large icon={<FunctionOutlined />} label="Formula Bar" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Zoom">
                <RibbonButton large icon={<SearchOutlined />} label="Zoom" onClick={openNativeFeature} native />
                <RibbonButton large icon={<NumberOutlined />} label="100%" onClick={() => spreadsheet?.setZoom(1)} disabled={!ready} />
                <RibbonButton large icon={<SearchOutlined />} label="Zoom to Selection" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Window">
                <RibbonButton large icon={<WindowsOutlined />} label="New Window" onClick={openNativeFeature} native />
                <RibbonButton large icon={<LayoutOutlined />} label="Freeze Panes" onClick={() => toggle('freeze')} disabled={editingDisabled} />
                <RibbonButton large icon={<LayoutOutlined />} label="Split" onClick={openNativeFeature} native />
            </RibbonGroup>
            <RibbonGroup label="Macros">
                <RibbonButton large icon={<WindowsOutlined />} label="Native VBE" onClick={onOpenVbe} />
                <RibbonButton large icon={<CodeOutlined />} label="VS Code VBA Studio" onClick={onOpenVbaDeveloper} />
            </RibbonGroup>
        </>
    );

    const renderAiVbaTab = () => (
        <>
            <RibbonGroup label="AI Context">
                <RibbonButton large icon={<RobotOutlined />} label="Export Workbook Context" onClick={onExportWorkbookContext} />
                <RibbonButton large icon={<CopyOutlined />} label="Ask Copilot" onClick={onAskCopilotAboutWorkbook} />
            </RibbonGroup>
            <RibbonGroup label="VBA">
                <RibbonButton large icon={<FileExcelOutlined />} label="Open in Excel" onClick={onOpenExcel} />
                <RibbonButton large icon={<WindowsOutlined />} label="Open native VBE" onClick={onOpenVbe} />
                <RibbonButton large icon={<CodeOutlined />} label="Open VBA Studio" onClick={onOpenVbaDeveloper} />
                <RibbonButton large icon={<FileTextOutlined />} label="Reveal VBA Sources" onClick={onOpenVbaExplorer} />
            </RibbonGroup>
            <RibbonGroup label="Workbook Tools">
                <RibbonButton large icon={<ColumnHeightOutlined />} label="AutoFit Columns" onClick={onAutoFitColumns} disabled={!ready} />
                <RibbonButton large icon={<SearchOutlined />} label="Find & Replace" onClick={() => command('find')} disabled={!ready} />
                <RibbonButton large icon={<PrinterOutlined />} label="Print" onClick={() => command('print')} disabled={!ready} />
            </RibbonGroup>
        </>
    );

    const renderActiveTab = () => {
        switch (activeTab) {
            case 'file':
                return renderFileTab();
            case 'insert':
                return renderInsertTab();
            case 'layout':
                return renderLayoutTab();
            case 'formulas':
                return renderFormulasTab();
            case 'data':
                return renderDataTab();
            case 'review':
                return renderReviewTab();
            case 'view':
                return renderViewTab();
            case 'ai-vba':
                return renderAiVbaTab();
            case 'home':
            default:
                return renderHomeTab();
        }
    };

    return (
        <>
        <div className="excel-ribbon" data-readonly={readOnly || undefined}>
            <div className="excel-ribbon-tabs-row">
                <div className="excel-ribbon-brand" title="Excel AI & VBA Studio">
                    <FileExcelOutlined />
                </div>
                <div className="excel-ribbon-tabs" role="tablist" aria-label="Excel ribbon">
                    {tabs.map((tab) => (
                        <button
                            key={tab.id}
                            type="button"
                            role="tab"
                            aria-selected={activeTab === tab.id}
                            className={`excel-ribbon-tab${activeTab === tab.id ? ' is-active' : ''}${tab.id === 'file' ? ' is-file' : ''}${tab.id === 'ai-vba' ? ' is-ai' : ''}`}
                            onClick={() => setActiveTab(tab.id)}
                        >
                            {tab.label}
                        </button>
                    ))}
                </div>
                <div className="excel-ribbon-quick" aria-label="Quick access">
                    <QuickButton icon={<SaveOutlined />} label="Save" onClick={() => command('save')} disabled={editingDisabled} />
                    <QuickButton icon={<UndoOutlined />} label="Undo" onClick={() => command('undo')} disabled={editingDisabled} />
                    <QuickButton icon={<RedoOutlined />} label="Redo" onClick={() => command('redo')} disabled={editingDisabled} />
                </div>
            </div>
            <div className="excel-ribbon-content" role="tabpanel">
                {renderActiveTab()}
            </div>
        </div>
        <Modal
            open={conditionalOpen}
            title="Mise en forme conditionnelle"
            onCancel={() => setConditionalOpen(false)}
            onOk={applyConditionalFormatting}
            okText="Appliquer à la sélection"
            cancelText="Annuler"
            width={460}
            footer={(_, { OkBtn, CancelBtn }) => (
                <>
                    <Button
                        danger
                        onClick={() => {
                            spreadsheet?.clearConditionalFormatting();
                            setConditionalOpen(false);
                        }}
                    >
                        Effacer les règles de la feuille
                    </Button>
                    <span style={{ flex: 1 }} />
                    <CancelBtn />
                    <OkBtn />
                </>
            )}
        >
            <div className="excel-conditional-dialog">
                <label>
                    <span>Type de règle</span>
                    <Select
                        value={conditionalPreset}
                        onChange={(value) => setConditionalPreset(value as ConditionalPreset)}
                        options={[
                            { value: 'greaterThan', label: 'Valeur supérieure à' },
                            { value: 'lessThan', label: 'Valeur inférieure à' },
                            { value: 'equal', label: 'Valeur égale à' },
                            { value: 'containsText', label: 'Texte contenant' },
                            { value: 'colorScale', label: 'Échelle de trois couleurs' },
                            { value: 'dataBar', label: 'Barres de données' },
                            { value: 'iconSet', label: 'Jeu de trois icônes' },
                        ]}
                    />
                </label>
                {!['colorScale', 'dataBar', 'iconSet'].includes(conditionalPreset) && (
                    <label>
                        <span>Valeur ou texte</span>
                        <Input
                            value={conditionalValue}
                            onChange={(event) => setConditionalValue(event.target.value)}
                            onPressEnter={applyConditionalFormatting}
                        />
                    </label>
                )}
                <p>
                    La règle sera appliquée aux cellules actuellement sélectionnées et enregistrée dans le fichier XLSX ou XLSM.
                </p>
            </div>
        </Modal>
        <TableDesigner
            open={tableDesignerOpen}
            tables={activeTables}
            selectionRangeRef={selectionRangeRef}
            inventoryAvailable={tableInventoryAvailable}
            isWorkbookTableNameAvailable={tableSpreadsheet?.isWorkbookTableNameAvailable?.bind(tableSpreadsheet)}
            readOnly={readOnly}
            onCancel={() => setTableDesignerOpen(false)}
            onCreate={createTable}
            onUpdate={updateTable}
            onDelete={removeTable}
        />
        <ChartDesigner
            open={chartDesignerOpen}
            charts={activeCharts}
            unsupportedChartCount={unsupportedNativeChartCount}
            selectionRangeRef={selectionRangeRef}
            initialChartType={chartInitialType}
            readOnly={readOnly}
            onCancel={() => setChartDesignerOpen(false)}
            onSave={saveChart}
            onDelete={removeChart}
        />
        <Modal
            open={integratedFeature !== null}
            title={integratedFeature ?? 'Outil Excel intégré'}
            onCancel={() => setIntegratedFeature(null)}
            onOk={() => void runIntegratedFeature()}
            okText="Exécuter dans VS Code"
            cancelText="Fermer"
            confirmLoading={featureRunning}
            width={520}
        >
            <div className="excel-integrated-feature-dialog">
                <p>
                    Cette fonction s’exécute dans l’éditeur intégré. Microsoft Excel ne sera pas ouvert.
                </p>
                {featureResult && (
                    <div className={featureResult.startsWith('Erreur') ? 'is-error' : 'is-success'}>
                        {featureResult}
                    </div>
                )}
            </div>
        </Modal>
        </>
    );
}
