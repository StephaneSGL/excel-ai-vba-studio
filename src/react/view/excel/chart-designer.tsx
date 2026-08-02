import {
    Alert,
    Button,
    Checkbox,
    Input,
    InputNumber,
    Modal,
    Select,
    Tabs,
} from 'antd';
import { useEffect, useMemo, useState } from 'react';
import type {
	ChartDataLabelPosition,
    ChartLegendPosition,
    ChartMarkerStyle,
    SheetChartAxisData,
    SheetChartData,
    SheetChartSeriesData,
} from '../../../common/excelWorkbookObjects';
import {
	chartTypeSupportsAxes,
	chartAxisGroupSupportsCategoryScale,
	chartDataLabelsHaveEnabledShowOption,
	chartDataLabelsHaveExplicitShowOption,
	canonicalChartTypeForSeries,
	chartSeriesSupportsBubbleSizes,
	chartSeriesSupportedDataLabelPositions,
	chartSeriesSupportsDataLabelPosition,
	chartSeriesSupportsPercentageDataLabels,
	chartSeriesSupportsSmooth,
    chartTypeSupportsGapWidth,
    chartTypeSupportsOverlap,
    chartTypeSupportsSecondaryAxes,
    isChartTypeCreatable,
} from '../../../common/excelWorkbookObjects';
import {
    CHART_TYPE_GROUPS,
    cloneChartDraft,
    createChartDraft,
    createSeriesDraft,
    findChartTypeLabel,
    validateAndNormalizeChart,
} from './chart-designer-model';
import './chart-designer.less';

type AxisKey = 'categoryAxis' | 'valueAxis' | 'secondaryCategoryAxis' | 'secondaryValueAxis';

function clearUnsupportedSeriesOptions(
    series: SheetChartSeriesData,
    inheritedChartType: number,
): SheetChartSeriesData {
    const effectiveChartType = series.chartType ?? inheritedChartType;
    const next = {
        ...series,
        ...(series.dataLabels ? { dataLabels: { ...series.dataLabels } } : {}),
    };
    if (!chartSeriesSupportsSmooth(effectiveChartType)) delete next.smooth;
    if (
        next.dataLabels?.showBubbleSize === true
        && !chartSeriesSupportsBubbleSizes(effectiveChartType)
    ) delete next.dataLabels.showBubbleSize;
    if (
        next.dataLabels?.showPercentage === true
        && !chartSeriesSupportsPercentageDataLabels(effectiveChartType)
    ) delete next.dataLabels.showPercentage;
	if (
		next.dataLabels?.position !== undefined
		&& !chartSeriesSupportsDataLabelPosition(
			effectiveChartType,
			next.dataLabels.position,
		)
	) delete next.dataLabels.position;
	if (
		next.dataLabels?.position !== undefined
		&& !chartDataLabelsHaveEnabledShowOption(next.dataLabels)
	) delete next.dataLabels.position;
	if (next.dataLabels && !chartDataLabelsHaveExplicitShowOption(next.dataLabels)) {
		delete next.dataLabels;
	}
    return next;
}

export interface ChartDesignerProps {
    open: boolean;
    charts: readonly SheetChartData[];
    unsupportedChartCount?: number;
    selectionRangeRef: string;
    initialChartType: number;
    readOnly?: boolean;
    onCancel: () => void;
    onSave: (chart: SheetChartData) => void;
    onDelete: (chartId: string) => void;
}

interface FieldProps {
    label: string;
    hint?: string;
    wide?: boolean;
    children: React.ReactNode;
}

interface AxisEditorProps {
    label: string;
    axis: SheetChartAxisData | undefined;
    defaultVisible?: boolean;
    supportsScale?: boolean;
    supportsGridlines?: boolean;
    onChange: (axis: SheetChartAxisData) => void;
}

const chartTypeOptions = CHART_TYPE_GROUPS.map(group => ({
    label: group.label,
    options: group.options.map(option => ({
        value: option.value,
        label: `${option.label} — ${option.constant}`,
        title: option.modern ? 'Type moderne : dépend de la version locale d’Excel' : option.constant,
        disabled: !isChartTypeCreatable(option.value),
    })),
}));

const legendOptions: { value: ChartLegendPosition; label: string }[] = [
    { value: 'right', label: 'Droite' },
    { value: 'left', label: 'Gauche' },
    { value: 'top', label: 'Haut' },
    { value: 'bottom', label: 'Bas' },
    { value: 'corner', label: 'Coin' },
    { value: 'custom', label: 'Position personnalisée Excel' },
];

const dataLabelPositionOptions: { value: ChartDataLabelPosition; label: string }[] = [
	{ value: 'bestFit', label: 'Meilleur ajustement' },
	{ value: 'above', label: 'Au-dessus' },
	{ value: 'below', label: 'En dessous' },
	{ value: 'center', label: 'Centre' },
	{ value: 'insideBase', label: 'Intérieur, base' },
	{ value: 'insideEnd', label: 'Intérieur, extrémité' },
	{ value: 'outsideEnd', label: 'Extérieur, extrémité' },
	{ value: 'left', label: 'Gauche' },
	{ value: 'right', label: 'Droite' },
];

const markerOptions: { value: ChartMarkerStyle; label: string }[] = [
    { value: 'automatic', label: 'Automatique' },
    { value: 'none', label: 'Aucun' },
    { value: 'circle', label: 'Cercle' },
    { value: 'square', label: 'Carré' },
    { value: 'diamond', label: 'Losange' },
    { value: 'triangle', label: 'Triangle' },
    { value: 'x', label: 'Croix X' },
    { value: 'plus', label: 'Plus' },
    { value: 'star', label: 'Étoile' },
    { value: 'dash', label: 'Tiret' },
    { value: 'dot', label: 'Point' },
    { value: 'picture', label: 'Image' },
];

function Field({ label, hint, wide = false, children }: FieldProps) {
    return (
        <label className={`chart-designer-field${wide ? ' is-wide' : ''}`}>
            <span className="chart-designer-field-label">{label}</span>
            {children}
            {hint && <span className="chart-designer-field-hint">{hint}</span>}
        </label>
    );
}

function ColorField({
    label,
    value,
    fallback,
    onChange,
}: {
    label: string;
    value: string | undefined;
    fallback: string;
    onChange: (value: string | undefined) => void;
}) {
    return (
        <div className="chart-designer-field">
            <span className="chart-designer-field-label">{label}</span>
            <div className="chart-designer-color-row">
                <input
                    type="color"
                    value={value || fallback}
                    aria-label={label}
                    onChange={event => onChange(event.target.value)}
                />
                <Input
                    value={value ?? ''}
                    placeholder="Automatique"
                    onChange={event => onChange(event.target.value || undefined)}
                />
                <Button size="small" onClick={() => onChange(undefined)}>Auto</Button>
            </div>
        </div>
    );
}

function AxisEditor({
    label,
    axis,
    defaultVisible = true,
    supportsScale = true,
    supportsGridlines = true,
    onChange,
}: AxisEditorProps) {
    const current: SheetChartAxisData = axis ?? { visible: defaultVisible };
    const update = <K extends keyof SheetChartAxisData>(key: K, value: SheetChartAxisData[K]) => {
        onChange({ ...current, [key]: value });
    };
    return (
        <fieldset className="chart-designer-axis">
            <legend>{label}</legend>
            <div className="chart-designer-check-row">
                <Checkbox checked={current.visible !== false} onChange={event => update('visible', event.target.checked)}>
                    Visible
                </Checkbox>
                {supportsScale && (
                    <Checkbox checked={Boolean(current.logarithmic)} onChange={event => update('logarithmic', event.target.checked)}>
                        Logarithmique
                    </Checkbox>
                )}
                <Checkbox checked={Boolean(current.reverseOrder)} onChange={event => update('reverseOrder', event.target.checked)}>
                    Ordre inversé
                </Checkbox>
                {supportsGridlines && (
                    <>
                        <Checkbox checked={Boolean(current.majorGridlines)} onChange={event => update('majorGridlines', event.target.checked)}>
                            Quadrillage principal
                        </Checkbox>
                        <Checkbox checked={Boolean(current.minorGridlines)} onChange={event => update('minorGridlines', event.target.checked)}>
                            Quadrillage secondaire
                        </Checkbox>
                    </>
                )}
            </div>
            <div className="chart-designer-grid is-four-columns">
                <Field label="Titre" wide>
                    <Input
                        value={current.title ?? ''}
                        maxLength={1000}
                        onChange={event => update('title', event.target.value)}
                    />
                </Field>
                <Field label="Format numérique" wide>
                    <Input
                        value={current.numberFormat ?? ''}
                        maxLength={255}
                        placeholder="Automatique, 0.00, 0 %, …"
                        onChange={event => update('numberFormat', event.target.value)}
                    />
                </Field>
                {supportsScale && <Field label="Minimum">
                    <InputNumber
                        value={current.minimumScale ?? null}
                        placeholder="Auto"
                        onChange={value => update('minimumScale', value)}
                    />
                </Field>}
                {supportsScale && <Field label="Maximum">
                    <InputNumber
                        value={current.maximumScale ?? null}
                        placeholder="Auto"
                        onChange={value => update('maximumScale', value)}
                    />
                </Field>}
                {supportsScale && <Field label="Unité principale">
                    <InputNumber
                        min={0}
                        value={current.majorUnit ?? null}
                        placeholder="Auto"
                        onChange={value => update('majorUnit', value)}
                    />
                </Field>}
                {supportsScale && <Field label="Unité secondaire">
                    <InputNumber
                        min={0}
                        value={current.minorUnit ?? null}
                        placeholder="Auto"
                        onChange={value => update('minorUnit', value)}
                    />
                </Field>}
            </div>
        </fieldset>
    );
}

export default function ChartDesigner({
    open,
    charts,
    unsupportedChartCount = 0,
    selectionRangeRef,
    initialChartType,
    readOnly = false,
    onCancel,
    onSave,
    onDelete,
}: ChartDesignerProps) {
    const [draft, setDraft] = useState<SheetChartData>(() => createChartDraft(charts, selectionRangeRef, initialChartType));
    const [selectedSeriesId, setSelectedSeriesId] = useState<string | null>(null);
    const [errors, setErrors] = useState<string[]>([]);

    useEffect(() => {
        if (!open) return;
        setDraft(createChartDraft(charts, selectionRangeRef, initialChartType));
        setSelectedSeriesId(null);
        setErrors([]);
    }, [open, initialChartType, selectionRangeRef]);

    const existingChart = charts.find(chart => chart.id === draft.id);
    const existing = Boolean(existingChart);
	const canPreserveCustomLegend = existingChart?.legend?.position === 'custom';
    const series = draft.series ?? [];
    const selectedSeries = series.find(item => item.id === selectedSeriesId) ?? null;
    const selectedSeriesIndex = series.findIndex(item => item.id === selectedSeriesId);
    const modernType = useMemo(() => (
        CHART_TYPE_GROUPS.flatMap(group => group.options).find(option => option.value === draft.chartType)?.modern === true
    ), [draft.chartType]);
    const explicitSeriesTypes = series.map(item => item.chartType ?? draft.chartType);
    const supportsAxes = chartTypeSupportsAxes(draft.chartType);
    const supportsSecondaryAxes = chartTypeSupportsSecondaryAxes(
        draft.chartType,
        explicitSeriesTypes,
    );
	const primaryCategoryAxisSupportsScale = chartAxisGroupSupportsCategoryScale(
		draft.chartType,
		series,
		'primary',
	);
	const secondaryCategoryAxisSupportsScale = chartAxisGroupSupportsCategoryScale(
		draft.chartType,
		series,
		'secondary',
	);
    const supportsGapWidth = chartTypeSupportsGapWidth(draft.chartType, explicitSeriesTypes);
    const supportsOverlap = chartTypeSupportsOverlap(draft.chartType, explicitSeriesTypes);
	const selectedSeriesSupportsBubbleSizes = selectedSeries
		? chartSeriesSupportsBubbleSizes(selectedSeries.chartType ?? draft.chartType)
		: false;
	const selectedSeriesSupportsSmooth = selectedSeries
		? chartSeriesSupportsSmooth(selectedSeries.chartType ?? draft.chartType)
		: false;
	const selectedSeriesSupportsPercentage = selectedSeries
		? chartSeriesSupportsPercentageDataLabels(selectedSeries.chartType ?? draft.chartType)
		: false;
	const selectedSeriesDataLabelPositions = selectedSeries
		? chartSeriesSupportedDataLabelPositions(selectedSeries.chartType ?? draft.chartType)
		: [];
	const selectedSeriesDataLabelPositionOptions = dataLabelPositionOptions.filter(
		option => selectedSeriesDataLabelPositions.includes(option.value),
	);

	const updateSeries = (next: SheetChartSeriesData) => {
		setDraft(current => {
			const compatibleNext = clearUnsupportedSeriesOptions(next, current.chartType);
			let nextSeries = (current.series ?? []).map(
				item => item.id === compatibleNext.id ? compatibleNext : item
			);
			const effectiveTypes = nextSeries.map(item => item.chartType ?? current.chartType);
			const nextChartType = canonicalChartTypeForSeries(
				current.chartType,
				effectiveTypes,
			);
			if (nextChartType === -4152 && current.chartType !== -4152) {
				nextSeries = nextSeries.map(item => ({
					...item,
					chartType: item.chartType ?? current.chartType,
				}));
			}
			return { ...current, chartType: nextChartType, series: nextSeries };
		});
    };

    const updateAxis = (key: AxisKey, axis: SheetChartAxisData) => {
        setDraft(current => ({ ...current, [key]: axis }));
    };

    const startNewChart = () => {
        setDraft(createChartDraft(charts, selectionRangeRef, initialChartType));
        setSelectedSeriesId(null);
        setErrors([]);
    };

    const selectChart = (chart: SheetChartData) => {
        const cloned = cloneChartDraft(chart);
		if (cloned.series) {
			cloned.series = cloned.series.map(seriesItem => (
				clearUnsupportedSeriesOptions(seriesItem, cloned.chartType)
			));
		}
        setDraft(cloned);
        setSelectedSeriesId(cloned.series?.[0]?.id ?? null);
        setErrors([]);
    };

    const addSeries = () => {
        const next = createSeriesDraft(draft.sourceRangeRef ?? selectionRangeRef);
        setDraft(current => ({
            ...current,
            sourceRangeRef: undefined,
            series: [...(current.series ?? []), next],
        }));
        setSelectedSeriesId(next.id);
    };

    const updateSourceRange = (sourceRangeRef: string) => {
        setDraft(current => ({ ...current, sourceRangeRef, series: undefined }));
        setSelectedSeriesId(null);
    };

    const removeSelectedSeries = () => {
        if (!selectedSeriesId) return;
        const remaining = series.filter(item => item.id !== selectedSeriesId);
        setDraft(current => ({ ...current, series: remaining }));
        setSelectedSeriesId(remaining[0]?.id ?? null);
    };

    const moveSelectedSeries = (offset: -1 | 1) => {
        if (selectedSeriesIndex < 0) return;
        const targetIndex = selectedSeriesIndex + offset;
        if (targetIndex < 0 || targetIndex >= series.length) return;
        setDraft(current => {
            const nextSeries = [...(current.series ?? [])];
            [nextSeries[selectedSeriesIndex], nextSeries[targetIndex]] = [
                nextSeries[targetIndex],
                nextSeries[selectedSeriesIndex],
            ];
            return { ...current, series: nextSeries };
        });
    };

    const save = () => {
        const result = validateAndNormalizeChart(draft, charts);
        if (!result.chart) {
            setErrors(result.errors);
            return;
        }
        try {
            onSave(result.chart);
            setErrors([]);
            onCancel();
        } catch (error) {
            setErrors([error instanceof Error ? error.message : String(error)]);
        }
    };

    const deleteChart = () => {
        if (!existing || !window.confirm(`Supprimer le graphique « ${draft.name} » ?`)) return;
        onDelete(draft.id);
        setDraft(createChartDraft(charts.filter(chart => chart.id !== draft.id), selectionRangeRef, initialChartType));
        setSelectedSeriesId(null);
        setErrors([]);
    };

    const generalTab = (
        <div className="chart-designer-section">
            <div className="chart-designer-grid is-four-columns">
                <Field label="Nom de l’objet" wide>
                    <Input value={draft.name} maxLength={255} onChange={event => setDraft(current => ({ ...current, name: event.target.value }))} />
                </Field>
                <Field label="Type Excel" wide>
                    <Select
                        showSearch
                        value={draft.chartType}
                        options={chartTypeOptions}
                        optionFilterProp="label"
                        popupMatchSelectWidth={520}
                        onChange={value => setDraft(current => {
							const nextSeries = current.series?.map(seriesItem => (
								clearUnsupportedSeriesOptions(seriesItem, value)
							));
							const nextSeriesTypes = nextSeries?.map(
								seriesItem => seriesItem.chartType ?? value
							) ?? [];
							return {
								...current,
								chartType: value,
								series: nextSeries,
								style: current.style ?? 2,
								gapWidth: chartTypeSupportsGapWidth(value, nextSeriesTypes)
									? current.gapWidth ?? 150
									: undefined,
								overlap: chartTypeSupportsOverlap(value, nextSeriesTypes)
									? current.overlap ?? 0
									: undefined,
							};
						})}
                    />
                </Field>
                <Field
                    label="Plage source"
                    hint="Plage A1 de la feuille active, par exemple A1:D20. Laissez vide uniquement avec des séries explicites."
                    wide
                >
                    <div className="chart-designer-inline">
                        <Input
                            value={draft.sourceRangeRef ?? ''}
                            placeholder="A1:D20"
                            onChange={event => updateSourceRange(event.target.value)}
                        />
                        <Button
                            disabled={!selectionRangeRef}
                            onClick={() => updateSourceRange(selectionRangeRef)}
                        >
                            Sélection actuelle
                        </Button>
                    </div>
                </Field>
                <Field label="Tracer les séries par">
                    <Select
                        value={draft.plotBy}
                        options={[
                            { value: 'columns', label: 'Colonnes' },
                            { value: 'rows', label: 'Lignes' },
                        ]}
                        onChange={value => setDraft(current => ({ ...current, plotBy: value }))}
                    />
                </Field>
                <Field label="Style Excel (1–48)">
                    <InputNumber min={1} max={48} value={draft.style} onChange={value => setDraft(current => ({ ...current, style: value ?? 2 }))} />
                </Field>
                {supportsGapWidth && <Field label="Intervalle (0–500)">
                    <InputNumber min={0} max={500} step={1} precision={0} value={draft.gapWidth} onChange={value => setDraft(current => ({ ...current, gapWidth: value ?? 150 }))} />
                </Field>}
                {supportsOverlap && <Field label="Chevauchement (-100–100)">
                    <InputNumber min={-100} max={100} step={1} precision={0} value={draft.overlap} onChange={value => setDraft(current => ({ ...current, overlap: value ?? 0 }))} />
                </Field>}
            </div>

            {modernType && (
                <Alert
                    type="info"
                    showIcon
                    message="Ce type moderne dépend de la version de Microsoft Excel installée."
                />
            )}

            <fieldset className="chart-designer-card">
                <legend>Titre et légende</legend>
                <div className="chart-designer-check-row">
                    <Checkbox
                        checked={Boolean(draft.title?.visible)}
                        onChange={event => setDraft(current => ({
                            ...current,
                            title: { visible: event.target.checked, text: current.title?.text ?? '' },
                        }))}
                    >
                        Afficher le titre
                    </Checkbox>
                    <Checkbox
                        checked={draft.legend?.visible !== false}
                        onChange={event => setDraft(current => ({
                            ...current,
                            legend: { visible: event.target.checked, position: current.legend?.position ?? 'right' },
                        }))}
                    >
                        Afficher la légende
                    </Checkbox>
                    <Checkbox
                        checked={Boolean(draft.roundedCorners)}
                        onChange={event => setDraft(current => ({ ...current, roundedCorners: event.target.checked }))}
                    >
                        Coins arrondis
                    </Checkbox>
                </div>
                <div className="chart-designer-grid is-two-columns">
                    <Field label="Texte du titre">
                        <Input
                            value={draft.title?.text ?? ''}
                            maxLength={1000}
                            onChange={event => setDraft(current => ({
                                ...current,
                                title: { visible: current.title?.visible ?? true, text: event.target.value },
                            }))}
                        />
                    </Field>
                    <Field label="Position de la légende">
                        <Select
                            value={draft.legend?.position ?? 'right'}
							options={legendOptions.map(option => ({
								...option,
								disabled: option.value === 'custom' && !canPreserveCustomLegend,
							}))}
                            onChange={value => setDraft(current => ({
                                ...current,
                                legend: { visible: current.legend?.visible !== false, position: value },
                            }))}
                        />
                    </Field>
                </div>
            </fieldset>

            <fieldset className="chart-designer-card">
                <legend>Position et dimensions (points Excel)</legend>
                <div className="chart-designer-grid is-four-columns">
                    {(['left', 'top', 'width', 'height'] as const).map(key => (
                        <Field key={key} label={{ left: 'Gauche', top: 'Haut', width: 'Largeur', height: 'Hauteur' }[key]}>
                            <InputNumber
                                min={key === 'width' || key === 'height' ? 20 : 0}
                                value={draft.anchor[key]}
                                onChange={value => setDraft(current => ({
                                    ...current,
                                    anchor: { ...current.anchor, [key]: value ?? 0 },
                                }))}
                            />
                        </Field>
                    ))}
                </div>
            </fieldset>

            <Field label="Texte alternatif" wide>
                <Input.TextArea
                    rows={2}
                    maxLength={1000}
                    value={draft.alternativeText ?? ''}
                    onChange={event => setDraft(current => ({ ...current, alternativeText: event.target.value }))}
                />
            </Field>
        </div>
    );

    const seriesTab = (
        <div className="chart-designer-series-layout">
            <aside className="chart-designer-series-list">
                <div className="chart-designer-list-actions">
                    <Button type="primary" size="small" onClick={addSeries}>Ajouter</Button>
                    <Button
                        size="small"
                        disabled={selectedSeriesIndex <= 0}
                        onClick={() => moveSelectedSeries(-1)}
                    >
                        Monter
                    </Button>
                    <Button
                        size="small"
                        disabled={selectedSeriesIndex < 0 || selectedSeriesIndex >= series.length - 1}
                        onClick={() => moveSelectedSeries(1)}
                    >
                        Descendre
                    </Button>
                    <Button size="small" danger disabled={!selectedSeries} onClick={removeSelectedSeries}>Supprimer</Button>
                </div>
                {series.length === 0 && (
                    <p className="chart-designer-empty">Aucune série explicite. Excel déduira les séries de la plage source.</p>
                )}
                {series.map((item, index) => (
                    <button
                        type="button"
                        key={item.id}
                        className={item.id === selectedSeriesId ? 'is-selected' : ''}
                        onClick={() => setSelectedSeriesId(item.id)}
                    >
                        <strong>{item.name?.trim() || `Série ${index + 1}`}</strong>
                        <span>{item.valuesRange || 'Plage à définir'}</span>
                    </button>
                ))}
            </aside>
            <section className="chart-designer-series-editor">
                {!selectedSeries && (
                    <Alert
                        type="info"
                        message="Ajoutez ou sélectionnez une série pour définir ses plages, son axe et son apparence."
                    />
                )}
                {selectedSeries && (
                    <>
                        <div className="chart-designer-check-row">
                            <Checkbox
                                checked={selectedSeries.visible !== false}
                                onChange={event => updateSeries({ ...selectedSeries, visible: event.target.checked })}
                            >
                                Série visible
                            </Checkbox>
                            <Checkbox
                                checked={selectedSeriesSupportsSmooth && Boolean(selectedSeries.smooth)}
								disabled={!selectedSeriesSupportsSmooth}
                                onChange={event => updateSeries({ ...selectedSeries, smooth: event.target.checked })}
                            >
                                Courbe lissée
                            </Checkbox>
                        </div>
                        <div className="chart-designer-grid is-two-columns">
                            <Field label="Nom libre">
                                <Input maxLength={255} value={selectedSeries.name ?? ''} onChange={event => updateSeries({ ...selectedSeries, name: event.target.value })} />
                            </Field>
                            <Field label="Cellule du nom">
                                <Input value={selectedSeries.nameRange ?? ''} placeholder="$B$1" onChange={event => updateSeries({ ...selectedSeries, nameRange: event.target.value })} />
                            </Field>
                            <Field label="Catégories">
                                <Input value={selectedSeries.categoryRange ?? ''} placeholder="$A$2:$A$20" onChange={event => updateSeries({ ...selectedSeries, categoryRange: event.target.value })} />
                            </Field>
                            <Field label="Valeurs Y (obligatoire)">
                                <Input value={selectedSeries.valuesRange} placeholder="$B$2:$B$20" onChange={event => updateSeries({ ...selectedSeries, valuesRange: event.target.value })} />
                            </Field>
                            <Field label="Valeurs X (nuage de points)">
                                <Input value={selectedSeries.xValuesRange ?? ''} placeholder="$A$2:$A$20" onChange={event => updateSeries({ ...selectedSeries, xValuesRange: event.target.value })} />
                            </Field>
                            <Field label="Tailles de bulles">
								<Input
									value={selectedSeries.bubbleSizesRange ?? ''}
									placeholder="$C$2:$C$20"
									disabled={!selectedSeriesSupportsBubbleSizes && !selectedSeries.bubbleSizesRange}
									onChange={event => updateSeries({ ...selectedSeries, bubbleSizesRange: event.target.value })}
								/>
                            </Field>
                            <Field label="Type de cette série" hint="Vide : hérite du type général.">
                                <Select
                                    allowClear
                                    showSearch
                                    value={selectedSeries.chartType}
                                    placeholder="Hériter du graphique"
                                    options={chartTypeOptions}
                                    optionFilterProp="label"
                                    popupMatchSelectWidth={520}
                                    onChange={value => updateSeries({ ...selectedSeries, chartType: value })}
                                />
                            </Field>
                            <Field label="Axe">
                                <Select
                                    value={selectedSeries.axisGroup ?? 'primary'}
                                    options={[
                                        { value: 'primary', label: 'Axe principal' },
                                        { value: 'secondary', label: 'Axe secondaire', disabled: !supportsSecondaryAxes },
                                    ]}
                                    onChange={value => updateSeries({ ...selectedSeries, axisGroup: value })}
                                />
                            </Field>
                        </div>

                        <fieldset className="chart-designer-card">
                            <legend>Couleurs, ligne et marqueurs</legend>
                            <div className="chart-designer-grid is-two-columns">
                                <ColorField label="Remplissage" value={selectedSeries.color} fallback="#4472c4" onChange={value => updateSeries({ ...selectedSeries, color: value })} />
                                <ColorField label="Couleur de ligne" value={selectedSeries.lineColor} fallback="#4472c4" onChange={value => updateSeries({ ...selectedSeries, lineColor: value })} />
                                <Field label="Épaisseur de ligne">
                                    <InputNumber min={0.1} max={20} step={0.25} value={selectedSeries.lineWidth} onChange={value => updateSeries({ ...selectedSeries, lineWidth: value ?? undefined })} />
                                </Field>
                                <Field label="Style de trait">
                                    <Select
                                        value={selectedSeries.dashStyle ?? 'solid'}
                                        options={[
                                            { value: 'solid', label: 'Continu' },
                                            { value: 'dash', label: 'Tirets' },
                                            { value: 'dot', label: 'Points' },
                                            { value: 'dashDot', label: 'Tiret-point' },
                                        ]}
                                        onChange={value => updateSeries({ ...selectedSeries, dashStyle: value })}
                                    />
                                </Field>
                                <Field label="Marqueur">
                                    <Select
                                        value={selectedSeries.markerStyle ?? 'automatic'}
                                        options={markerOptions}
                                        onChange={value => updateSeries({ ...selectedSeries, markerStyle: value })}
                                    />
                                </Field>
                                <Field label="Taille du marqueur">
                                    <InputNumber min={2} max={72} step={1} precision={0} value={selectedSeries.markerSize} onChange={value => updateSeries({ ...selectedSeries, markerSize: value ?? undefined })} />
                                </Field>
                            </div>
                        </fieldset>

                        <fieldset className="chart-designer-card">
                            <legend>Étiquettes de données</legend>
                            <div className="chart-designer-check-row">
                                {([
                                    ['showValue', 'Valeur'],
                                    ['showCategoryName', 'Catégorie'],
                                    ['showSeriesName', 'Nom de série'],
                                    ['showPercentage', 'Pourcentage'],
                                    ['showBubbleSize', 'Taille de bulle'],
                                ] as const).map(([key, label]) => (
                                    <Checkbox
                                        key={key}
										disabled={
											(key === 'showBubbleSize' && !selectedSeriesSupportsBubbleSizes)
											|| (key === 'showPercentage' && !selectedSeriesSupportsPercentage)
										}
										checked={Boolean(selectedSeries.dataLabels?.[key]) && !(
											(key === 'showBubbleSize' && !selectedSeriesSupportsBubbleSizes)
											|| (key === 'showPercentage' && !selectedSeriesSupportsPercentage)
										)}
                                        onChange={event => {
											const dataLabels = {
												...selectedSeries.dataLabels,
												[key]: event.target.checked,
											};
											if (!chartDataLabelsHaveEnabledShowOption(dataLabels)) {
												delete dataLabels.position;
											}
											updateSeries({ ...selectedSeries, dataLabels });
										}}
                                    >
                                        {label}
                                    </Checkbox>
                                ))}
                            </div>
                            <Field label="Position">
                                <Select
                                    value={selectedSeries.dataLabels?.position}
									disabled={selectedSeriesDataLabelPositionOptions.length === 0}
									allowClear
									placeholder="Position automatique"
                                    options={selectedSeriesDataLabelPositionOptions}
                                    onChange={value => updateSeries({
                                        ...selectedSeries,
                                        dataLabels: {
											...selectedSeries.dataLabels,
											...(chartDataLabelsHaveEnabledShowOption(selectedSeries.dataLabels)
												? {}
												: { showValue: true }),
											position: value,
										},
                                    })}
                                />
                            </Field>
                        </fieldset>
                    </>
                )}
            </section>
        </div>
    );

    const axesTab = (
        <div className="chart-designer-section">
			<AxisEditor label="Axe principal des catégories" axis={draft.categoryAxis} supportsScale={primaryCategoryAxisSupportsScale} onChange={axis => updateAxis('categoryAxis', axis)} />
            <AxisEditor label="Axe principal des valeurs" axis={draft.valueAxis} onChange={axis => updateAxis('valueAxis', axis)} />
			{supportsSecondaryAxes && <AxisEditor label="Axe secondaire des catégories" axis={draft.secondaryCategoryAxis} defaultVisible={false} supportsScale={secondaryCategoryAxisSupportsScale} supportsGridlines={false} onChange={axis => updateAxis('secondaryCategoryAxis', axis)} />}
            {supportsSecondaryAxes && <AxisEditor label="Axe secondaire des valeurs" axis={draft.secondaryValueAxis} defaultVisible={false} supportsGridlines={false} onChange={axis => updateAxis('secondaryValueAxis', axis)} />}
        </div>
    );

    return (
        <Modal
            open={open}
            title="Concepteur de graphiques Excel"
            width={1180}
            onCancel={onCancel}
            destroyOnClose
            footer={[
                <Button key="delete" danger disabled={!existing || readOnly} onClick={deleteChart} className="chart-designer-delete">
                    Supprimer le graphique
                </Button>,
                <Button key="cancel" onClick={onCancel}>Annuler</Button>,
                <Button key="save" type="primary" disabled={readOnly} onClick={save}>
                    {existing ? 'Mettre à jour' : 'Créer le graphique'}
                </Button>,
            ]}
        >
            <div className="chart-designer-shell">
                <aside className="chart-designer-chart-list" aria-label="Graphiques de la feuille">
                    <div className="chart-designer-list-heading">
                        <strong>Graphiques de la feuille</strong>
                        <Button size="small" onClick={startNewChart}>Nouveau</Button>
                    </div>
                    {charts.length === 0 && <p className="chart-designer-empty">Aucun graphique enregistré sur cette feuille.</p>}
                    {charts.map(chart => (
                        <button
                            type="button"
                            key={chart.id}
                            className={chart.id === draft.id ? 'is-selected' : ''}
                            onClick={() => selectChart(chart)}
                        >
                            <strong>{chart.name}</strong>
                            <span>{findChartTypeLabel(chart.chartType)}</span>
                            <span>{chart.sourceRangeRef || `${chart.series?.length ?? 0} série(s)`}</span>
                        </button>
                    ))}
                </aside>
                <main className="chart-designer-main">
                    {readOnly && <Alert type="warning" showIcon message="Classeur en lecture seule : consultation uniquement." />}
                    {unsupportedChartCount > 0 && (
                        <Alert
                            type="warning"
                            showIcon
                            message={`${unsupportedChartCount} graphique(s) natif(s) sont conservés mais ne peuvent pas être modifiés dans cet éditeur.`}
                        />
                    )}
                    {errors.length > 0 && (
                        <Alert
                            type="error"
                            showIcon
                            message="Le graphique ne peut pas être enregistré"
                            description={<ul>{errors.map(error => <li key={error}>{error}</li>)}</ul>}
                            closable
                            onClose={() => setErrors([])}
                        />
                    )}
                    <Tabs
                        defaultActiveKey="general"
                        items={[
                            { key: 'general', label: 'Général et disposition', children: generalTab },
                            { key: 'series', label: `Séries (${series.length})`, children: seriesTab },
                            ...(supportsAxes
                                ? [{ key: 'axes', label: 'Axes et quadrillages', children: axesTab }]
                                : []),
                        ]}
                    />
                </main>
            </div>
        </Modal>
    );
}
