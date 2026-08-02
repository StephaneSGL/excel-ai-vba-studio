import {
    EXCEL_CHART_TYPES,
    MAX_WORKBOOK_OBJECT_RANGE_CELLS,
    SIMPLE_A1_RANGE,
    chartTypeSupportsAxes,
    chartTypeSupportsSecondaryAxes,
	chartAxisGroupSupportsCategoryScale,
	chartSeriesSupportsBubbleSizes,
	chartSeriesSupportsDataLabelPosition,
	chartDataLabelsHaveEnabledShowOption,
	chartDataLabelsHaveExplicitShowOption,
	chartSeriesSupportsPercentageDataLabels,
	chartSeriesSupportsSmooth,
	chartSeriesTypesCanCoexist,
	canonicalChartTypeForSeries,
    chartTypeSupportsGapWidth,
    chartTypeSupportsOverlap,
    isChartTypeCreatable,
    parseSimpleA1Range,
    type ExcelChartTypeOption,
    type SheetChartAxisData,
    type SheetChartData,
    type SheetChartSeriesData,
} from '../../../common/excelWorkbookObjects';

export const DEFAULT_CHART_TYPE = 51;

export interface ChartTypeGroup {
    label: string;
    options: ExcelChartTypeOption[];
}

export interface ChartValidationResult {
    chart: SheetChartData | null;
    errors: string[];
}

export const CHART_TYPE_GROUPS: ChartTypeGroup[] = Array.from(
    EXCEL_CHART_TYPES.reduce((groups, option) => {
        const entries = groups.get(option.group) ?? [];
        entries.push(option);
        groups.set(option.group, entries);
        return groups;
    }, new Map<string, ExcelChartTypeOption[]>()),
    ([label, options]) => ({ label, options }),
);

const HEX_COLOR = /^#[0-9a-f]{6}$/i;

function createStableId(prefix: string): string {
    const randomUuid = globalThis.crypto?.randomUUID?.();
    if (randomUuid) return `${prefix}-${randomUuid}`;
    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function clone<T>(value: T): T {
    return JSON.parse(JSON.stringify(value)) as T;
}

function trimOptional(value: string | undefined): string | undefined {
    const normalized = value?.trim();
    return normalized ? normalized : undefined;
}

function finiteOrNull(value: number | null | undefined): number | null {
    return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function normalizeAxis(
    axis: SheetChartAxisData,
    defaultVisible = true,
    supportsScale = true,
    supportsGridlines = true,
): SheetChartAxisData {
    const normalized: SheetChartAxisData = {
        visible: axis?.visible ?? defaultVisible,
		// Axis definitions are full modeled UI state. An explicit empty title
		// tells the native bridge to set HasTitle=false instead of preserving a
		// title the user just cleared.
        title: trimOptional(axis?.title) ?? '',
        reverseOrder: Boolean(axis?.reverseOrder),
        // Empty means "follow the source". It is deliberately serialized so
        // clearing a custom format is not mistaken for "leave unchanged".
        numberFormat: trimOptional(axis?.numberFormat) ?? '',
    };
    if (supportsScale) {
        normalized.minimumScale = finiteOrNull(axis?.minimumScale);
        normalized.maximumScale = finiteOrNull(axis?.maximumScale);
        normalized.majorUnit = finiteOrNull(axis?.majorUnit);
        normalized.minorUnit = finiteOrNull(axis?.minorUnit);
        normalized.logarithmic = Boolean(axis?.logarithmic);
    }
    if (supportsGridlines) {
        normalized.majorGridlines = Boolean(axis?.majorGridlines);
        normalized.minorGridlines = Boolean(axis?.minorGridlines);
    }
    return normalized;
}

function validateAxis(
    label: string,
    axis: SheetChartAxisData | undefined,
    errors: string[],
    supportsScale = true,
): void {
    if (!axis) return;
    if (supportsScale) {
        const minimum = finiteOrNull(axis.minimumScale);
        const maximum = finiteOrNull(axis.maximumScale);
        const majorUnit = finiteOrNull(axis.majorUnit);
        const minorUnit = finiteOrNull(axis.minorUnit);
        if (axis.minimumScale != null && minimum == null) errors.push(`${label} : le minimum doit être un nombre fini.`);
        if (axis.maximumScale != null && maximum == null) errors.push(`${label} : le maximum doit être un nombre fini.`);
        if (minimum != null && maximum != null && minimum >= maximum) {
            errors.push(`${label} : le minimum doit être inférieur au maximum.`);
        }
        if (axis.majorUnit != null && (majorUnit == null || majorUnit <= 0)) {
            errors.push(`${label} : l’unité principale doit être strictement positive.`);
        }
        if (axis.minorUnit != null && (minorUnit == null || minorUnit <= 0)) {
            errors.push(`${label} : l’unité secondaire doit être strictement positive.`);
        }
        if (axis.logarithmic && (
            (minimum != null && minimum <= 0)
            || (maximum != null && maximum <= 0)
        )) {
            errors.push(`${label} : une échelle logarithmique exige des bornes positives.`);
        }
	} else if (
		axis.minimumScale != null
		|| axis.maximumScale != null
		|| axis.majorUnit != null
		|| axis.minorUnit != null
		|| axis.logarithmic === true
	) {
		errors.push(`${label} : les bornes numériques exigent une série nuage de points ou bulles sur ce groupe d’axes.`);
    }
    if ((axis.title ?? '').length > 1000) {
        errors.push(`${label} : le titre ne peut pas dépasser 1 000 caractères.`);
    }
    if ((axis.numberFormat ?? '').length > 255) {
        errors.push(`${label} : le format numérique ne peut pas dépasser 255 caractères.`);
    }
}

function validateRange(label: string, value: string | undefined, errors: string[], required = false): void {
    const normalized = value?.trim() ?? '';
    if (!normalized) {
        if (required) errors.push(`${label} est obligatoire.`);
        return;
    }
    if (!isValidWorkbookRangeRef(normalized)) {
        errors.push(`${label} n’est pas une plage Excel locale valide.`);
    } else if ((parseSimpleA1Range(normalized)?.cellCount ?? Infinity) > MAX_WORKBOOK_OBJECT_RANGE_CELLS) {
        errors.push(`${label} dépasse la limite de ${MAX_WORKBOOK_OBJECT_RANGE_CELLS.toLocaleString('fr-FR')} cellules.`);
    }
}

function normalizeSeries(series: SheetChartSeriesData): SheetChartSeriesData {
    const normalized: SheetChartSeriesData = {
        id: series.id,
        valuesRange: series.valuesRange.trim(),
    };
    for (const key of ['name', 'nameRange', 'categoryRange', 'xValuesRange', 'bubbleSizesRange', 'color', 'lineColor'] as const) {
        const value = trimOptional(series[key]);
        if (value !== undefined) normalized[key] = value;
    }
    for (const key of ['chartType', 'lineWidth', 'markerSize'] as const) {
        if (typeof series[key] === 'number') normalized[key] = series[key];
    }
    for (const key of ['axisGroup', 'dashStyle', 'markerStyle'] as const) {
        if (series[key] !== undefined) normalized[key] = series[key] as never;
    }
    for (const key of ['smooth', 'visible'] as const) {
        if (series[key] !== undefined) normalized[key] = Boolean(series[key]);
    }
    if (series.dataLabels !== undefined) {
        const labels: NonNullable<SheetChartSeriesData['dataLabels']> = {};
        for (const key of ['showValue', 'showCategoryName', 'showSeriesName', 'showPercentage', 'showBubbleSize'] as const) {
            if (series.dataLabels[key] !== undefined) labels[key] = Boolean(series.dataLabels[key]);
        }
        if (series.dataLabels.position !== undefined) labels.position = series.dataLabels.position;
        if (chartDataLabelsHaveExplicitShowOption(labels)) normalized.dataLabels = labels;
    }
    return normalized;
}

export function findChartTypeLabel(value: number | undefined): string {
    if (value == null) return 'Type du graphique';
    const option = EXCEL_CHART_TYPES.find(candidate => candidate.value === value);
    return option ? `${option.label} (${option.constant})` : `Type Excel ${value}`;
}

export function isValidWorkbookRangeRef(value: string): boolean {
    const normalized = value.trim();
    return SIMPLE_A1_RANGE.test(normalized.toUpperCase()) && parseSimpleA1Range(normalized) !== null;
}

export function createSeriesDraft(sourceRangeRef = ''): SheetChartSeriesData {
    return {
        id: createStableId('series'),
        name: '',
        nameRange: '',
        categoryRange: '',
        valuesRange: sourceRangeRef,
        xValuesRange: '',
        bubbleSizesRange: '',
    };
}

export function createChartDraft(
    charts: readonly SheetChartData[],
    sourceRangeRef = '',
    initialChartType = DEFAULT_CHART_TYPE,
): SheetChartData {
    const usedNames = new Set(charts.map(chart => chart.name.trim().toLocaleLowerCase()));
    let sequence = charts.length + 1;
    while (usedNames.has(`chart ${sequence}`)) sequence += 1;
    return {
        id: createStableId('chart'),
        name: `Chart ${sequence}`,
        chartType: initialChartType,
        sourceRangeRef: sourceRangeRef.trim() || undefined,
        plotBy: 'columns',
        anchor: { left: 40, top: 80, width: 640, height: 360 },
        title: { visible: false, text: '' },
        legend: { visible: true, position: 'right' },
        ...(chartTypeSupportsAxes(initialChartType) ? { categoryAxis: {
            visible: true,
            minimumScale: null,
            maximumScale: null,
            majorUnit: null,
            minorUnit: null,
            majorGridlines: false,
            minorGridlines: false,
        } } : {}),
        ...(chartTypeSupportsAxes(initialChartType) ? { valueAxis: {
            visible: true,
            minimumScale: null,
            maximumScale: null,
            majorUnit: null,
            minorUnit: null,
            majorGridlines: true,
            minorGridlines: false,
        } } : {}),
        ...(chartTypeSupportsSecondaryAxes(initialChartType) ? {
            secondaryCategoryAxis: { visible: false },
            secondaryValueAxis: { visible: false },
        } : {}),
        style: 2,
        roundedCorners: false,
        ...(chartTypeSupportsGapWidth(initialChartType) ? { gapWidth: 150 } : {}),
        ...(chartTypeSupportsOverlap(initialChartType) ? { overlap: 0 } : {}),
        alternativeText: '',
    };
}

export function cloneChartDraft(chart: SheetChartData): SheetChartData {
    return clone(chart);
}

export function validateAndNormalizeChart(
    draft: SheetChartData,
    existingCharts: readonly SheetChartData[],
): ChartValidationResult {
    const errors: string[] = [];
    const name = draft.name.trim();
    if (!name) errors.push('Le nom du graphique est obligatoire.');
    if (name.length > 255) errors.push('Le nom du graphique ne peut pas dépasser 255 caractères.');
    if (name.includes('\0')) errors.push('Le nom du graphique contient un caractère interdit.');
    if (existingCharts.some(chart => chart.id !== draft.id && chart.name.trim().toLocaleLowerCase() === name.toLocaleLowerCase())) {
        errors.push('Un autre graphique de la feuille porte déjà ce nom.');
    }
    if (!EXCEL_CHART_TYPES.some(option => option.value === draft.chartType)) {
        errors.push('Le type de graphique sélectionné ne fait pas partie de l’inventaire Excel pris en charge.');
    } else if (!isChartTypeCreatable(draft.chartType)) {
        errors.push(draft.chartType === 140
            ? 'Le graphique cartographique est désactivé : Excel peut transmettre ses données à Bing Maps pendant sa création.'
            : 'Le type « Graphique recommandé » est consultatif et ne peut pas être enregistré comme type concret.');
    }
    validateRange('La plage source', draft.sourceRangeRef, errors);

    const { left, top, width, height } = draft.anchor;
    if (![left, top, width, height].every(Number.isFinite)) errors.push('La position et les dimensions doivent être des nombres finis.');
    if (left < 0 || top < 0) errors.push('La position du graphique ne peut pas être négative.');
    if (width < 20 || height < 20) {
        errors.push('La largeur et la hauteur doivent être supérieures ou égales à 20 points.');
    }
    if (left > 1_000_000 || top > 1_000_000 || width > 100_000 || height > 100_000) {
        errors.push('La position ou la taille du graphique dépasse les limites de sécurité.');
    }
    if (draft.style != null && (!Number.isInteger(draft.style) || draft.style < 1 || draft.style > 48)) {
        errors.push('Le style doit être un entier compris entre 1 et 48.');
    }
    if (draft.gapWidth != null && (!Number.isInteger(draft.gapWidth) || draft.gapWidth < 0 || draft.gapWidth > 500)) {
        errors.push('La largeur d’intervalle doit être un entier compris entre 0 et 500.');
    }
    if (draft.overlap != null && (!Number.isInteger(draft.overlap) || draft.overlap < -100 || draft.overlap > 100)) {
        errors.push('Le chevauchement doit être un entier compris entre -100 et 100.');
    }
    if ((draft.title?.text ?? '').length > 1000) errors.push('Le titre ne peut pas dépasser 1 000 caractères.');
    if ((draft.alternativeText ?? '').length > 1000) errors.push('Le texte alternatif ne peut pas dépasser 1 000 caractères.');

    const series = draft.series ?? [];
	const existingChart = existingCharts.find(chart => chart.id === draft.id);
	if (
		draft.legend?.position === 'custom'
		&& existingChart?.legend?.position !== 'custom'
	) {
		errors.push('La position personnalisée de la légende peut seulement conserver une disposition manuelle déjà créée dans Excel.');
	}
    const seriesChartTypes = series.map(item => item.chartType ?? draft.chartType);
	if (!chartSeriesTypesCanCoexist(seriesChartTypes)) {
		errors.push('Un graphique ne peut pas mélanger des séries à bulles et des séries d’un autre type : Excel convertit alors silencieusement toutes les séries en bulles.');
	}
	const effectiveChartType = canonicalChartTypeForSeries(
		draft.chartType,
		seriesChartTypes,
	);
	const normalizedSeries = effectiveChartType === -4152 && draft.chartType !== -4152
		? series.map(item => ({ ...item, chartType: item.chartType ?? draft.chartType }))
		: series;
	const primaryCategoryAxisSupportsScale = chartAxisGroupSupportsCategoryScale(
		effectiveChartType,
		normalizedSeries,
		'primary',
	);
	const secondaryCategoryAxisSupportsScale = chartAxisGroupSupportsCategoryScale(
		effectiveChartType,
		normalizedSeries,
		'secondary',
	);
    if (!draft.sourceRangeRef?.trim() && series.length === 0) {
        errors.push('Indiquez une plage source ou ajoutez au moins une série explicite.');
    }
    if (series.length > 255) errors.push('Un graphique ne peut pas contenir plus de 255 séries dans cet éditeur.');
    const seriesIds = new Set<string>();
    series.forEach((item, index) => {
        const label = `Série ${index + 1}`;
		const effectiveSeriesType = item.chartType ?? draft.chartType;
        if (!item.id || seriesIds.has(item.id)) errors.push(`${label} possède un identifiant manquant ou dupliqué.`);
        seriesIds.add(item.id);
        validateRange(`${label} — valeurs`, item.valuesRange, errors, true);
        validateRange(`${label} — nom`, item.nameRange, errors);
        const nameRangeBounds = item.nameRange?.trim()
            ? parseSimpleA1Range(item.nameRange.trim())
            : null;
        if (nameRangeBounds && nameRangeBounds.cellCount !== 1) {
            errors.push(`${label} : la plage du nom doit désigner une seule cellule.`);
        }
        validateRange(`${label} — catégories`, item.categoryRange, errors);
        validateRange(`${label} — valeurs X`, item.xValuesRange, errors);
        validateRange(`${label} — tailles de bulles`, item.bubbleSizesRange, errors);
		if (item.bubbleSizesRange?.trim() && !chartSeriesSupportsBubbleSizes(effectiveSeriesType)) {
			errors.push(`${label} : la plage de tailles exige un type de graphique à bulles.`);
		}
		if (item.smooth !== undefined && !chartSeriesSupportsSmooth(effectiveSeriesType)) {
			errors.push(`${label} : le lissage est réservé aux séries ligne ou nuage de points.`);
		}
		if (item.dataLabels?.showBubbleSize === true && !chartSeriesSupportsBubbleSizes(effectiveSeriesType)) {
			errors.push(`${label} : l’étiquette de taille de bulle exige une série de type bulle.`);
		}
		if (item.dataLabels !== undefined && !chartDataLabelsHaveExplicitShowOption(item.dataLabels)) {
			errors.push(`${label} : les étiquettes exigent au moins une option d’affichage explicite.`);
		}
		if (
			item.dataLabels?.position !== undefined
			&& !chartDataLabelsHaveEnabledShowOption(item.dataLabels)
		) {
			errors.push(`${label} : une position d’étiquette exige au moins une option d’affichage activée.`);
		}
		if (
			item.dataLabels?.showPercentage === true
			&& !chartSeriesSupportsPercentageDataLabels(effectiveSeriesType)
		) {
			errors.push(`${label} : l’étiquette de pourcentage exige une série de type secteur ou anneau.`);
		}
		if (
			item.dataLabels?.position !== undefined
			&& !chartSeriesSupportsDataLabelPosition(
				effectiveSeriesType,
				item.dataLabels.position,
			)
		) {
			errors.push(`${label} : la position d’étiquette ${item.dataLabels.position} n’est pas prise en charge par ce type de graphique.`);
		}
        if (item.categoryRange?.trim() && item.xValuesRange?.trim()) {
            errors.push(`${label} : catégories et valeurs X ne peuvent pas être définies ensemble.`);
        }
        if (item.name?.trim() && item.nameRange?.trim()) {
            errors.push(`${label} : le nom libre et la cellule du nom ne peuvent pas être définis ensemble.`);
        }
        if (item.name && /^[=+\-@]/.test(item.name.trim())) {
            errors.push(`${label} : le nom libre doit être du texte littéral, pas une formule.`);
        }
        if ((item.name ?? '').length > 255) {
            errors.push(`${label} : le nom libre ne peut pas dépasser 255 caractères.`);
        }
        if (item.chartType != null && !EXCEL_CHART_TYPES.some(option => option.value === item.chartType)) {
            errors.push(`${label} utilise un type de graphique inconnu.`);
        } else if (item.chartType != null && !isChartTypeCreatable(item.chartType)) {
            errors.push(`${label} doit utiliser un type de graphique local autorisé.`);
        }
        if (
            item.axisGroup === 'secondary' &&
            !chartTypeSupportsSecondaryAxes(effectiveSeriesType)
        ) {
            errors.push(`${label} ne peut pas utiliser un axe secondaire avec ce type de graphique.`);
        }
        if (item.lineWidth != null && (!Number.isFinite(item.lineWidth) || item.lineWidth < 0.1 || item.lineWidth > 20)) {
            errors.push(`${label} : l’épaisseur de ligne doit être comprise entre 0,1 et 20.`);
        }
        if (item.markerSize != null && (!Number.isInteger(item.markerSize) || item.markerSize < 2 || item.markerSize > 72)) {
            errors.push(`${label} : la taille de marqueur doit être un entier compris entre 2 et 72.`);
        }
        if (item.color?.trim() && !HEX_COLOR.test(item.color.trim())) {
            errors.push(`${label} : la couleur de remplissage doit utiliser le format #RRGGBB.`);
        }
        if (item.lineColor?.trim() && !HEX_COLOR.test(item.lineColor.trim())) {
            errors.push(`${label} : la couleur de ligne doit utiliser le format #RRGGBB.`);
        }
    });

    if (effectiveChartType === -4152 && (
        normalizedSeries.length === 0
        || normalizedSeries.some(item =>
            item.chartType === undefined ||
            item.chartType === -4152 ||
            !isChartTypeCreatable(item.chartType)
        )
        || new Set(normalizedSeries.map(item => item.chartType)).size < 2
    )) {
        errors.push('Un graphique combiné personnalisé exige au moins deux séries de types concrets distincts.');
    }
	if (chartTypeSupportsAxes(effectiveChartType)) {
		validateAxis('Axe des catégories', draft.categoryAxis, errors, primaryCategoryAxisSupportsScale);
        validateAxis('Axe des valeurs', draft.valueAxis, errors);
		if (chartTypeSupportsSecondaryAxes(effectiveChartType, seriesChartTypes)) {
			validateAxis('Axe secondaire des catégories', draft.secondaryCategoryAxis, errors, secondaryCategoryAxisSupportsScale);
            validateAxis('Axe secondaire des valeurs', draft.secondaryValueAxis, errors);
        }
    }
    if (errors.length) return { chart: null, errors };

    return {
        chart: {
            id: draft.id,
            name,
			chartType: effectiveChartType,
            ...(series.length === 0 && trimOptional(draft.sourceRangeRef) !== undefined
                ? { sourceRangeRef: trimOptional(draft.sourceRangeRef) }
                : {}),
            plotBy: draft.plotBy,
            anchor: { left, top, width, height },
            ...(draft.title === undefined ? {} : { title: {
                visible: Boolean(draft.title?.visible),
                text: draft.title?.text?.trim() ?? '',
            } }),
            ...(draft.legend === undefined ? {} : { legend: {
                visible: draft.legend?.visible !== false,
                position: draft.legend?.position ?? 'right',
            } }),
			...(chartTypeSupportsAxes(effectiveChartType) && draft.categoryAxis
				? { categoryAxis: normalizeAxis(draft.categoryAxis, true, primaryCategoryAxisSupportsScale, true) }
                : {}),
			...(chartTypeSupportsAxes(effectiveChartType) && draft.valueAxis
                ? { valueAxis: normalizeAxis(draft.valueAxis, true, true, true) }
                : {}),
			...(chartTypeSupportsSecondaryAxes(effectiveChartType, seriesChartTypes) && draft.secondaryCategoryAxis
				? { secondaryCategoryAxis: normalizeAxis(draft.secondaryCategoryAxis, false, secondaryCategoryAxisSupportsScale, false) }
                : {}),
			...(chartTypeSupportsSecondaryAxes(effectiveChartType, seriesChartTypes) && draft.secondaryValueAxis
                ? { secondaryValueAxis: normalizeAxis(draft.secondaryValueAxis, false, true, false) }
                : {}),
			...(normalizedSeries.length ? { series: normalizedSeries.map(normalizeSeries) } : {}),
            ...(draft.style === undefined ? {} : { style: draft.style }),
            ...(draft.roundedCorners === undefined ? {} : { roundedCorners: Boolean(draft.roundedCorners) }),
            ...(chartTypeSupportsGapWidth(
				effectiveChartType,
                seriesChartTypes,
            ) && draft.gapWidth !== undefined ? { gapWidth: draft.gapWidth } : {}),
            ...(chartTypeSupportsOverlap(
				effectiveChartType,
                seriesChartTypes,
            ) && draft.overlap !== undefined ? { overlap: draft.overlap } : {}),
			// The designer owns this modeled property. Keep an explicit empty
			// string so clearing existing alternative text reaches native Excel
			// instead of being misread as an omitted/preserve instruction.
			alternativeText: trimOptional(draft.alternativeText) ?? '',
        },
        errors: [],
    };
}
