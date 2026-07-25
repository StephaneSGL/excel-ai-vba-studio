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
import { useEffect, useState, type ReactNode } from 'react';
import type Spreadsheet from './x-spreadsheet/index';
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
    onOpenVbaDeveloper: () => void;
    onExportWorkbookContext: () => void;
    onOpenVbaExplorer: () => void;
    onAskCopilotAboutWorkbook: () => void;
};

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
                ? `${label} — intégration VS Code en préparation (Excel ne sera pas ouvert)`
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
            onClick={onClick}
            disabled={isDisabled}
            title={tooltip}
        >
            <span className="excel-ribbon-button-icon" aria-hidden>{icon}</span>
            <span className="excel-ribbon-button-label">{label}</span>
            {native && <span className="excel-ribbon-native-badge">Bientôt</span>}
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
    onOpenVbaDeveloper,
    onExportWorkbookContext,
    onOpenVbaExplorer,
    onAskCopilotAboutWorkbook,
}: ExcelRibbonProps) {
    const [activeTab, setActiveTab] = useState<RibbonTab>('home');
    const [gridVisible, setGridVisibleState] = useState(true);
    const [sheetProtected, setSheetProtected] = useState(false);
    const ready = spreadsheet !== null;
    const editingDisabled = !ready || readOnly;

    useEffect(() => {
        setGridVisibleState(true);
        setSheetProtected(spreadsheet?.isSheetProtected() ?? false);
    }, [spreadsheet]);

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
    const openNativeFeature = () => {
        window.alert(
            'Cette fonction restera dans VS Code. Son interface intégrée est encore en préparation ; Microsoft Excel ne sera pas ouvert.',
        );
    };
    const addConditionalFormatting = () => {
        if (!spreadsheet) return;
        const comparison = window.prompt(
            'Conditional formatting rule: enter >, >=, <, <=, =, or contains',
            '>',
        )?.trim().toLowerCase();
        if (!comparison) return;
        const expected = window.prompt('Value or text to compare with');
        if (expected == null) return;
        const operatorMap: Record<string, string> = {
            '>': 'greaterThan',
            '>=': 'greaterThanOrEqual',
            '<': 'lessThan',
            '<=': 'lessThanOrEqual',
            '=': 'equal',
            '==': 'equal',
        };
        const isContains = comparison === 'contains';
        const numericValue = Number(expected.replace(',', '.'));
        const formulaValue = Number.isNaN(numericValue) ? expected : numericValue;
        spreadsheet.addConditionalFormatting({
            type: isContains ? 'containsText' : 'cellIs',
            ...(isContains
                ? { operator: 'containsText', text: expected, formulae: [expected] }
                : {
                    operator: operatorMap[comparison] ?? 'greaterThan',
                    formulae: [formulaValue],
                }),
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
            },
        });
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
                <RibbonButton icon={<TableOutlined />} label="Format as Table" onClick={openNativeFeature} native />
                <RibbonButton icon={<BarChartOutlined />} label="Conditional Formatting" onClick={addConditionalFormatting} disabled={editingDisabled} />
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
                <RibbonButton large icon={<TableOutlined />} label="Table" onClick={openNativeFeature} native />
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
                <RibbonButton large icon={<BarChartOutlined />} label="Recommended Charts" onClick={openNativeFeature} native />
                <RibbonButton large icon={<BarChartOutlined />} label="Column Chart" onClick={openNativeFeature} native />
                <RibbonButton large icon={<BarChartOutlined />} label="Line Chart" onClick={openNativeFeature} native />
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
                <RibbonButton large icon={<CodeOutlined />} label="Macros / VBA" onClick={onOpenVbaDeveloper} />
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
    );
}
