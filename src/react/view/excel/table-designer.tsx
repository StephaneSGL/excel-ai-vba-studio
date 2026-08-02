import { Alert, Button, Checkbox, Input, Modal, Select } from 'antd';
import { useEffect, useMemo, useState } from 'react';
import {
    buildExcelTableStyleCatalog,
    excelTableNameComparisonKey,
    isValidExcelTableName,
    MAX_WORKBOOK_OBJECT_RANGE_CELLS,
	minimumExcelTableRangeRows,
    normalizeA1Range,
    normalizeExcelTableName,
    parseSimpleA1Range,
    SIMPLE_A1_RANGE,
    type SheetTableData,
    type SheetTableStyle,
} from '../../../common/excelWorkbookObjects';
import './table-designer.less';

export interface TableDesignerValues {
    name?: string;
    displayName?: string;
    rangeRef?: string;
    headerRow: boolean;
    totalsRow: boolean;
    style: SheetTableStyle;
}

export interface TableDesignerProps {
    open: boolean;
    tables: readonly SheetTableData[];
    selectionRangeRef: string;
    inventoryAvailable: boolean;
    readOnly?: boolean;
    isWorkbookTableNameAvailable?: (name: string, currentTableId?: string) => boolean;
    onCancel: () => void;
    onCreate: (values: TableDesignerValues) => void;
    onUpdate?: (table: SheetTableData) => void;
    onDelete?: (tableId: string) => void;
}

interface TableDraft extends TableDesignerValues {
    id?: string;
    nameText: string;
    displayNameText: string;
    rangeRefText: string;
}

const TABLE_STYLES = buildExcelTableStyleCatalog();

const styleOptions = ['Light', 'Medium', 'Dark'].map(group => ({
    label: `${group} (${TABLE_STYLES.filter(style => style.startsWith(`TableStyle${group}`)).length})`,
    options: TABLE_STYLES
        .filter(style => style.startsWith(`TableStyle${group}`))
        .map(style => ({ value: style, label: style.replace('TableStyle', '') })),
}));

function createDraft(selectionRangeRef: string): TableDraft {
    return {
        nameText: '',
        displayNameText: '',
        rangeRefText: normalizeA1Range(selectionRangeRef || ''),
        headerRow: true,
        totalsRow: false,
        style: {
            name: 'TableStyleMedium2',
            showFirstColumn: false,
            showLastColumn: false,
            showRowStripes: true,
            showColumnStripes: false,
        },
    };
}

function draftFromTable(table: SheetTableData): TableDraft {
    return {
        id: table.id,
        name: table.name,
        displayName: table.displayName,
        rangeRef: table.rangeRef,
        nameText: table.name,
        displayNameText: table.displayName,
        rangeRefText: table.rangeRef,
        headerRow: table.headerRow,
        totalsRow: table.totalsRow,
        style: { ...table.style },
    };
}

function validateDraft(
    draft: TableDraft,
    tables: readonly SheetTableData[],
    isWorkbookTableNameAvailable?: (name: string, currentTableId?: string) => boolean,
): string[] {
    const errors: string[] = [];
    const normalizedRange = normalizeA1Range(draft.rangeRefText);
	const editingTable = draft.id
		? tables.find(table => table.id === draft.id) ?? null
		: null;
	if (!editingTable && draft.totalsRow) {
		errors.push('La création native d’une ligne de totaux est refusée car Excel déplace les cellules et réécrit les références de formules.');
	}
	if (editingTable && draft.totalsRow !== editingTable.totalsRow) {
		errors.push('L’état de la ligne de totaux d’une table existante ne peut pas être modifié en sécurité.');
	}
	if (editingTable?.totalsRow && normalizedRange !== normalizeA1Range(editingTable.rangeRef)) {
		errors.push('La plage d’une table qui possède déjà une ligne de totaux ne peut pas être redimensionnée en sécurité.');
	}
    if (!SIMPLE_A1_RANGE.test(normalizedRange)) {
        errors.push('La plage doit être locale à la feuille, par exemple A1:D20.');
    } else {
        const parsedRange = parseSimpleA1Range(normalizedRange);
        if (!parsedRange) {
            errors.push('La plage dépasse les limites de lignes ou colonnes d’Excel.');
        } else if (parsedRange.cellCount > MAX_WORKBOOK_OBJECT_RANGE_CELLS) {
            errors.push(`La plage dépasse la limite de ${MAX_WORKBOOK_OBJECT_RANGE_CELLS.toLocaleString('fr-FR')} cellules.`);
		} else if (
			parsedRange.endRow - parsedRange.startRow + 1
			< minimumExcelTableRangeRows(draft.totalsRow)
		) {
			errors.push(draft.totalsRow
				? 'La plage doit contenir au moins trois lignes : en-tête, données et totaux.'
				: 'La plage doit contenir au moins deux lignes : en-tête et données.');
        }
    }
    const names = [draft.nameText, draft.displayNameText]
        .map(normalizeExcelTableName)
        .filter(Boolean);
    names.forEach(name => {
        if (!isValidExcelTableName(name)) {
            errors.push(`« ${name} » n’est pas un nom de tableau Excel valide.`);
        }
    });
    const normalizedName = normalizeExcelTableName(draft.nameText);
    const normalizedDisplayName = normalizeExcelTableName(draft.displayNameText);
    if (normalizedName && normalizedDisplayName
        && excelTableNameComparisonKey(normalizedName) !== excelTableNameComparisonKey(normalizedDisplayName)) {
        errors.push('Pour Excel, Name et DisplayName doivent rester identiques.');
    }
    const requestedName = normalizedName || normalizedDisplayName;
    const localNameAvailable = !tables.some(table => (
        table.id !== draft.id
        && [table.name, table.displayName].some(
            name => excelTableNameComparisonKey(name) === excelTableNameComparisonKey(requestedName)
        )
    ));
    if (requestedName && (
        isWorkbookTableNameAvailable
            ? !isWorkbookTableNameAvailable(requestedName, draft.id)
            : !localNameAvailable
    )) {
        errors.push('Une autre table du classeur utilise déjà ce nom.');
    }
    if (!TABLE_STYLES.includes(draft.style.name)) errors.push('Le style sélectionné ne fait pas partie des 60 styles Excel pris en charge.');
    return [...new Set(errors)];
}

function valuesFromDraft(draft: TableDraft): TableDesignerValues {
    const name = normalizeExcelTableName(draft.nameText) || undefined;
    const displayName = normalizeExcelTableName(draft.displayNameText) || name;
    return {
        name,
        displayName,
        rangeRef: normalizeA1Range(draft.rangeRefText),
        headerRow: draft.headerRow,
        totalsRow: draft.totalsRow,
        style: { ...draft.style },
    };
}

export default function TableDesigner({
    open,
    tables,
    selectionRangeRef,
    inventoryAvailable,
    readOnly = false,
    isWorkbookTableNameAvailable,
    onCancel,
    onCreate,
    onUpdate,
    onDelete,
}: TableDesignerProps) {
    const [draft, setDraft] = useState<TableDraft>(() => createDraft(selectionRangeRef));
    const [errors, setErrors] = useState<string[]>([]);
    const editingTable = draft.id ? tables.find(table => table.id === draft.id) ?? null : null;
    const canUpdate = Boolean(editingTable && onUpdate);
    const canDelete = Boolean(editingTable && onDelete);

    useEffect(() => {
        if (!open) return;
        setDraft(createDraft(selectionRangeRef));
        setErrors([]);
    }, [open, selectionRangeRef]);

    const styleGroup = useMemo(() => {
        if (draft.style.name.startsWith('TableStyleLight')) return 'light';
        if (draft.style.name.startsWith('TableStyleDark')) return 'dark';
        return 'medium';
    }, [draft.style.name]);

    const updateStyle = <K extends keyof SheetTableStyle>(key: K, value: SheetTableStyle[K]) => {
        setDraft(current => ({ ...current, style: { ...current.style, [key]: value } }));
    };

    const save = () => {
        const validationErrors = validateDraft(
            draft,
            tables,
            isWorkbookTableNameAvailable,
        );
        if (validationErrors.length) {
            setErrors(validationErrors);
            return;
        }
        try {
            const values = valuesFromDraft(draft);
            if (editingTable) {
                if (!onUpdate) {
                    setErrors(['La version actuelle du moteur permet d’inventorier cette table, mais pas encore de la modifier.']);
                    return;
                }
                onUpdate({
                    ...editingTable,
                    ...values,
                    name: values.name ?? editingTable.name,
                    displayName: values.displayName ?? values.name ?? editingTable.displayName,
                    rangeRef: values.rangeRef ?? editingTable.rangeRef,
                });
            } else {
                onCreate(values);
            }
            setErrors([]);
            onCancel();
        } catch (error) {
            setErrors([error instanceof Error ? error.message : String(error)]);
        }
    };

    const remove = () => {
        if (!editingTable || !onDelete || !window.confirm(`Supprimer la table « ${editingTable.name} » ?`)) return;
        try {
            onDelete(editingTable.id);
            setDraft(createDraft(selectionRangeRef));
            setErrors([]);
        } catch (error) {
            setErrors([error instanceof Error ? error.message : String(error)]);
        }
    };

    return (
        <Modal
            open={open}
            title="Concepteur de tableaux Excel"
            width={850}
            onCancel={onCancel}
            destroyOnClose
            footer={[
                <Button key="delete" danger className="table-designer-delete" disabled={!canDelete || readOnly} onClick={remove}>
                    Supprimer la table
                </Button>,
                <Button key="cancel" onClick={onCancel}>Annuler</Button>,
                <Button key="save" type="primary" disabled={readOnly || Boolean(editingTable && !canUpdate)} onClick={save}>
                    {editingTable ? 'Mettre à jour' : 'Créer la table'}
                </Button>,
            ]}
        >
            <div className="table-designer-shell">
                <aside className="table-designer-list" aria-label="Tables de la feuille">
                    <div className="table-designer-list-heading">
                        <strong>Tables de la feuille</strong>
                        <Button size="small" onClick={() => { setDraft(createDraft(selectionRangeRef)); setErrors([]); }}>
                            Nouvelle
                        </Button>
                    </div>
                    {!inventoryAvailable && (
                        <p className="table-designer-empty">L’inventaire détaillé n’est pas exposé par cette version du moteur.</p>
                    )}
                    {inventoryAvailable && tables.length === 0 && (
                        <p className="table-designer-empty">Aucune table sur cette feuille.</p>
                    )}
                    {tables.map(table => (
                        <button
                            type="button"
                            key={table.id}
                            className={draft.id === table.id ? 'is-selected' : ''}
                            onClick={() => { setDraft(draftFromTable(table)); setErrors([]); }}
                        >
                            <strong>{table.displayName || table.name}</strong>
                            <span>{table.rangeRef}</span>
                            <span>{table.style.name}</span>
                        </button>
                    ))}
                </aside>

                <main className="table-designer-main">
                    {readOnly && <Alert type="warning" showIcon message="Classeur en lecture seule : consultation uniquement." />}
                    {editingTable && (!onUpdate || !onDelete) && (
                        <Alert
                            type="info"
                            showIcon
                            message="Capacités limitées du moteur"
                            description={`Inventaire disponible. Modification : ${onUpdate ? 'oui' : 'non'} ; suppression : ${onDelete ? 'oui' : 'non'}.`}
                        />
                    )}
                    {errors.length > 0 && (
                        <Alert
                            type="error"
                            showIcon
                            closable
                            message="La table ne peut pas être enregistrée"
                            description={<ul>{errors.map(error => <li key={error}>{error}</li>)}</ul>}
                            onClose={() => setErrors([])}
                        />
                    )}

                    <div className="table-designer-grid">
                        <label>
                            <span>Nom</span>
                            <Input
                                value={draft.nameText}
                                placeholder="Table_Ventes"
                                maxLength={255}
                                onChange={event => setDraft(current => ({ ...current, nameText: event.target.value }))}
                            />
                        </label>
                        <label>
                            <span>DisplayName</span>
                            <Input
                                value={draft.displayNameText}
                                placeholder="Identique au nom si vide"
                                maxLength={255}
                                onChange={event => setDraft(current => ({ ...current, displayNameText: event.target.value }))}
                            />
                        </label>
                        <label className="is-wide">
                            <span>Plage de la table</span>
                            <div className="table-designer-inline">
                                <Input
                                    value={draft.rangeRefText}
                                    placeholder="A1:D20"
									disabled={editingTable?.totalsRow === true}
                                    onChange={event => setDraft(current => ({ ...current, rangeRefText: event.target.value }))}
                                />
                                <Button
									disabled={!selectionRangeRef || editingTable?.totalsRow === true}
                                    onClick={() => setDraft(current => ({ ...current, rangeRefText: normalizeA1Range(selectionRangeRef) }))}
                                >
                                    Sélection actuelle
                                </Button>
                            </div>
                            <small>Chaque plage disjointe peut devenir sa propre table, même dans les mêmes colonnes.</small>
                        </label>
                        <label className="is-wide">
                            <span>Style ({TABLE_STYLES.length} styles Excel)</span>
                            <Select
                                showSearch
                                value={draft.style.name}
                                options={styleOptions}
                                optionFilterProp="label"
                                popupMatchSelectWidth={360}
                                onChange={value => updateStyle('name', value)}
                            />
                        </label>
                    </div>

                    <div className={`table-designer-preview is-${styleGroup}`} aria-label={`Aperçu ${draft.style.name}`}>
                        <div>Colonne 1</div><div>Colonne 2</div><div>Colonne 3</div>
                        <div>Valeur A</div><div>120</div><div>Actif</div>
                        <div>Valeur B</div><div>240</div><div>Actif</div>
                    </div>

                    <div className="table-designer-options">
						<Checkbox
							checked={draft.headerRow}
							disabled
							onChange={event => setDraft(current => ({ ...current, headerRow: event.target.checked }))}
						>
                            Ligne d’en-tête
                        </Checkbox>
						<Checkbox
							checked={draft.totalsRow}
							disabled
							onChange={event => setDraft(current => ({ ...current, totalsRow: event.target.checked }))}
                        >
                            Ligne des totaux
                        </Checkbox>
						<small>Les totaux existants sont préservés; leur activation, désactivation et redimensionnement sont refusés pour protéger les cellules et formules.</small>
                        <Checkbox checked={draft.style.showFirstColumn} onChange={event => updateStyle('showFirstColumn', event.target.checked)}>
                            Première colonne accentuée
                        </Checkbox>
                        <Checkbox checked={draft.style.showLastColumn} onChange={event => updateStyle('showLastColumn', event.target.checked)}>
                            Dernière colonne accentuée
                        </Checkbox>
                        <Checkbox checked={draft.style.showRowStripes} onChange={event => updateStyle('showRowStripes', event.target.checked)}>
                            Lignes à bandes
                        </Checkbox>
                        <Checkbox checked={draft.style.showColumnStripes} onChange={event => updateStyle('showColumnStripes', event.target.checked)}>
                            Colonnes à bandes
                        </Checkbox>
                    </div>
                </main>
            </div>
        </Modal>
    );
}
